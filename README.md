# Word-to-PDF Sintegra

Aplicacao web simples para converter documentos entre Word (`.docx`) e PDF diretamente no navegador. O usuario faz o upload do arquivo, escolhe o formato de saida e acompanha o progresso de envio, conversao e download em tempo real. Os arquivos enviados ficam em `uploads/` e os convertidos em `converted/`, evitando sobrescrever o original.

## Pre-requisitos
- Windows 10 ou 11 com PowerShell 5.1 (ou superior) habilitado.
- Microsoft Word instalado (o script usa a automacao COM do Word para converter os arquivos).

## Como rodar localmente
1. Abra um terminal PowerShell e navegue ate a pasta do projeto.
2. Execute o servidor (altere a porta se necessario):
   ```powershell
   powershell.exe -ExecutionPolicy Bypass -File .\server.ps1 -Port 3000
   ```
3. Abra `http://localhost:3000` no navegador e envie um arquivo `.docx` ou `.pdf` para converter.
4. Apos a conversao, clique em **Baixar arquivo convertido** para receber o resultado.

> Dica: use `Ctrl+C` no terminal para encerrar o servidor. Se a porta estiver ocupada, basta iniciar com `-Port 3001` (ou outro numero livre).
