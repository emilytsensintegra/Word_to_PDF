param(
    [int]$Port = 3000
)

Add-Type -AssemblyName System.Web | Out-Null

function Initialize-Directories {
    $paths = @(
        "$PSScriptRoot\public",
        "$PSScriptRoot\uploads",
        "$PSScriptRoot\converted"
    )
    foreach ($p in $paths) { if (-not (Test-Path $p)) { New-Item -ItemType Directory -Path $p | Out-Null } }
}


function Get-ContentType([string]$path) {
    switch ([IO.Path]::GetExtension($path).ToLowerInvariant()) {
        '.html' { 'text/html; charset=utf-8'; break }
        '.css'  { 'text/css; charset=utf-8'; break }
        '.js'   { 'application/javascript; charset=utf-8'; break }
        '.png'  { 'image/png'; break }
        '.jpg'  { 'image/jpeg'; break }
        '.jpeg' { 'image/jpeg'; break }
        '.svg'  { 'image/svg+xml'; break }
        '.pdf'  { 'application/pdf'; break }
        '.docx' { 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'; break }
        default { 'application/octet-stream' }
    }
}

function Sanitize-Filename([string]$name) {
    # Keep only safe chars
    $safe = ($name -replace '[^A-Za-z0-9\._-]', '_')
    # Prevent empty
    if ([string]::IsNullOrWhiteSpace($safe)) { $safe = "upload" }
    return $safe
}

function Convert-Document {
    param(
        [Parameter(Mandatory)] [string]$InputPath,
        [Parameter(Mandatory)] [string]$OutputPath,
        [Parameter(Mandatory)] [ValidateSet('pdf','docx')] [string]$Target
    )

    if (-not (Test-Path $InputPath)) { throw "Input not found: $InputPath" }

    $wdFormatPDF = 17
    $wdFormatXMLDocument = 12  # .docx

    $word = $null
    $doc = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0

        $doc = $word.Documents.Open($InputPath)
        switch ($Target) {
            'pdf'  { $doc.SaveAs([ref]$OutputPath, [ref]$wdFormatPDF) }
            'docx' { $doc.SaveAs([ref]$OutputPath, [ref]$wdFormatXMLDocument) }
        }
    }
    finally {
        if ($doc -ne $null) { $doc.Close($false) | Out-Null }
        if ($word -ne $null) { $word.Quit() | Out-Null }
        if ($doc) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) }
        if ($word) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

function Read-RequestBodyToFile {
    param(
        [Parameter(Mandatory)] $Request,
        [Parameter(Mandatory)] [string]$DestPath
    )
    $fs = [System.IO.File]::Open($DestPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
    try {
        $Request.InputStream.CopyTo($fs)
    }
    finally { $fs.Close() }
}

function Write-ResponseBytes {
    param(
        [Parameter(Mandatory)] $Response,
        [Parameter(Mandatory)] [byte[]]$Bytes,
        [Parameter()] [int]$StatusCode = 200,
        [Parameter()] [string]$ContentType = 'application/octet-stream',
        [Parameter()] [hashtable]$Headers
    )
    $Response.StatusCode = $StatusCode
    $Response.ContentType = $ContentType
    if ($Headers) {
        foreach ($k in $Headers.Keys) { $Response.Headers.Add($k, [string]$Headers[$k]) }
    }
    $Response.ContentLength64 = $Bytes.LongLength
    $Response.OutputStream.Write($Bytes, 0, $Bytes.Length)
    $Response.OutputStream.Close()
    try { $Response.Close() } catch {}
}

function Write-ResponseText {
    param(
        [Parameter(Mandatory)] $Response,
        [Parameter(Mandatory)] [string]$Text,
        [Parameter()] [int]$StatusCode = 200,
        [Parameter()] [string]$ContentType = 'application/json; charset=utf-8',
        [Parameter()] [hashtable]$Headers
    )
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
    Write-ResponseBytes -Response $Response -Bytes $bytes -StatusCode $StatusCode -ContentType $ContentType -Headers $Headers
}

function Handle-Request {
    param($Context)

    $req = $Context.Request
    $res = $Context.Response

    $path = $req.Url.AbsolutePath
    if ($path -eq '/') { $path = '/index.html' }

    if ($req.HttpMethod -eq 'GET') {
        $publicRoot = Join-Path $PSScriptRoot 'public'
        $fsPath = Join-Path $publicRoot ($path.TrimStart('/'))
        if ((Test-Path $fsPath) -and -not (Get-Item $fsPath).PSIsContainer) {
            try {
                $bytes = [IO.File]::ReadAllBytes($fsPath)
                $ct = Get-ContentType -path $fsPath
                Write-ResponseBytes -Response $res -Bytes $bytes -ContentType $ct
            } catch {
                Write-ResponseText -Response $res -Text '{"error":"Failed to read file"}' -StatusCode 500
            }
            return
        } else {
            Write-ResponseText -Response $res -Text '{"error":"Not found"}' -StatusCode 404
            return
        }
    }

    if ($req.HttpMethod -eq 'POST' -and $path -eq '/convert') {
        try {
            $qs = [System.Web.HttpUtility]::ParseQueryString($req.Url.Query)
            # PowerShell 5.1 compatibility: replace null-coalescing (??) with explicit checks
            $target = $qs['target']
            if ($null -eq $target) { $target = '' }
            $target = $target.ToLowerInvariant()
            $filename = Sanitize-Filename($qs['filename'])
            Write-Host "[convert] start target=$target filename=$filename" -ForegroundColor Cyan
            if ([string]::IsNullOrWhiteSpace($target) -or ($target -notin @('pdf','docx'))) {
                Write-ResponseText -Response $res -Text '{"error":"Invalid target; use pdf or docx"}' -StatusCode 400
                return
            }
            if ([string]::IsNullOrWhiteSpace($filename)) { $filename = 'upload' }

            $uploads = Join-Path $PSScriptRoot 'uploads'
            $converted = Join-Path $PSScriptRoot 'converted'

            $inExt = [IO.Path]::GetExtension($filename)
            if ([string]::IsNullOrEmpty($inExt)) {
                # try infer from Content-Type
                $ct = $req.ContentType
                switch -regex ($ct) {
                    'pdf' { $inExt = '.pdf' }
                    'word' { $inExt = '.docx' }
                }
                if (-not $inExt) { $inExt = '.bin' }
                $filename = "$filename$inExt"
            }

            $inputPath = Join-Path $uploads $filename
            Read-RequestBodyToFile -Request $req -DestPath $inputPath
            try { $inSize = (Get-Item $inputPath).Length } catch { $inSize = 0 }
            Write-Host "[convert] saved upload to: $inputPath ($inSize bytes)" -ForegroundColor DarkCyan

            $outName = [IO.Path]::GetFileNameWithoutExtension($filename)
            $outExt = if ($target -eq 'pdf') { '.pdf' } else { '.docx' }
            $outputPath = Join-Path $converted ("$outName$outExt")

            # If source and target are same ext, just copy
            if ([IO.Path]::GetExtension($inputPath).ToLowerInvariant() -eq $outExt.ToLowerInvariant()) {
                Write-Host "[convert] same extension, copying to $outputPath" -ForegroundColor Yellow
                Copy-Item -Force -Path $inputPath -Destination $outputPath
            } else {
                Write-Host "[convert] converting via Word to $outExt..." -ForegroundColor Yellow
                Convert-Document -InputPath $inputPath -OutputPath $outputPath -Target $target
                Write-Host "[convert] conversion finished" -ForegroundColor Green
            }

            $bytes = [IO.File]::ReadAllBytes($outputPath)
            $ct = Get-ContentType -path $outputPath
            # Build Content-Disposition safely for PS 5.1 (no escape backslashes)
            $disp = 'attachment; filename="' + [IO.Path]::GetFileName($outputPath) + '"'
            Write-Host "[convert] sending response ($($bytes.Length) bytes) as $ct" -ForegroundColor DarkGreen
            Write-ResponseBytes -Response $res -Bytes $bytes -ContentType $ct -Headers @{ 'Content-Disposition' = $disp }
        }
        catch {
            $msg = ($_ | Out-String).Trim()
            Write-ResponseText -Response $res -Text (ConvertTo-Json @{ error = $msg }) -StatusCode 500
        }
        return
    }

    Write-ResponseText -Response $res -Text '{"error":"Not found"}' -StatusCode 404
}

function Start-Server {
    param([int]$Port)

    Initialize-Directories

    $listener = New-Object System.Net.HttpListener
    $prefix = "http://localhost:$Port/"
    $listener.Prefixes.Add($prefix)
    try {
        $listener.Start()
    } catch {
        Write-Error "Failed to start listener on $prefix. Try running PowerShell as Administrator or change the port with -Port. Error: $_"
        return
    }

    Write-Host "Server running at $prefix" -ForegroundColor Green
    Write-Host "Open $prefix in your browser. Press Ctrl+C to stop." -ForegroundColor Yellow

    try {
        while ($listener.IsListening) {
            $context = $listener.GetContext()
            Handle-Request -Context $context
        }
    }
    finally {
        $listener.Stop()
        $listener.Close()
    }
}

Start-Server -Port $Port
