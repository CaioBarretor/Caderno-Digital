# Caderno Faculdade

Aplicativo Windows (Python/Tkinter) e primeira versão de cliente web instalável para Android.

## Windows
Baixe a versão portátil na página Releases. O aplicativo Windows atual ainda grava DOCX na pasta local escolhida; usando uma pasta do Google Drive para computador, ela é sincronizada pelo cliente oficial.

## Android / Web
O cliente PWA está na pasta web/. Ele foi preparado para autenticar com OAuth 2.0, usar o Google Picker para escolher uma pasta e ler/gravar DOCX pela Google Drive API v3. A aplicação só funciona depois de configurar as credenciais Google em web/config.js e publicar em HTTPS.

Veja web/CONFIGURAR-GOOGLE.md para configurar o projeto Google Cloud. Nunca coloque Client Secret em código de navegador.

## Limitação atual de sincronização
O cliente Android usa a Drive API diretamente. O cliente Windows ainda não foi migrado para a API; por enquanto, a ponte entre os dois é a pasta local sincronizada pelo Google Drive para computador. Para uma sincronização direta e previsível entre ambos, a próxima etapa é migrar também o cliente Windows para OAuth/Drive API e testar a permissão de acesso aos documentos existentes.
