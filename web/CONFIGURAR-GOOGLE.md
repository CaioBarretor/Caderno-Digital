# Configurar Google Drive para o Caderno

1. No Google Cloud Console, crie um projeto e ative Google Drive API e Google Picker API.
2. Configure a tela de consentimento OAuth. Enquanto estiver em modo de teste, adicione sua conta Google como usuário de teste.
3. Crie uma chave de API e um OAuth Client ID do tipo Web application.
4. No OAuth Client ID, adicione a origem HTTPS onde a aplicação será hospedada. Restrinja a chave de API à origem do site e às APIs necessárias.
5. Copie o número do projeto Google Cloud.
6. Edite web/config.js e preencha GOOGLE_CLIENT_ID, GOOGLE_API_KEY e GOOGLE_APP_ID.
7. Publique a pasta web/ em hospedagem HTTPS. Abra o site no Chrome do Android e use o menu do navegador > Adicionar à tela inicial.

Não coloque Client Secret no navegador ou no repositório. A chave de API e o Client ID são identificadores públicos, mas devem ter restrições adequadas.

A aplicação usa OAuth 2.0 com o escopo drive.file e o seletor oficial Google Picker. Esse escopo dá acesso a arquivos criados pelo app ou explicitamente abertos/selecionados com ele; documentos antigos criados fora do app podem precisar ser selecionados uma vez. A sincronização direta do aplicativo Windows com a mesma API ainda é uma etapa separada; enquanto isso, o cliente Windows continua usando a pasta local do Drive para computador.
