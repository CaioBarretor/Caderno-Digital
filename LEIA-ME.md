# Caderno Faculdade

Aplicativo desktop local para organizar anotações por período e disciplina e salvar cada matéria como arquivo `.docx`.

## Recursos
- Interface de caderno com título e editor de texto.
- Organização automática: `Pasta escolhida / Período / Disciplina / Matéria.docx`.
- Navegação por período, disciplina e busca pelo nome.
- Abre os DOCX criados pelo aplicativo.
- Tema claro e escuro.
- Guarda a pasta e o tema escolhidos entre execuções.
- Os arquivos são documentos Word normais, acessíveis fora do aplicativo.

## Sincronização com Google Drive
Esta versão usa a sincronização do **Google Drive para computador**: escolha no aplicativo uma pasta que já esteja dentro da área sincronizada do Drive. Ao salvar uma matéria, o app grava o `.docx` nessa pasta e o aplicativo oficial do Drive sincroniza o arquivo.

Não é necessário inserir senha Google no aplicativo nem configurar credenciais de API. A sincronização depende do Google Drive para computador estar instalado, conectado e com a pasta escolhida sincronizada. O app não faz sincronização direta via API e não contorna conflitos/offline do Drive.

## Executar no Windows
1. Instale Python 3.11 ou superior, marcando a opção de adicionar Python ao PATH.
2. Extraia esta pasta para um local permanente.
3. Dê dois cliques em `Executar Caderno.bat`. Na primeira execução, ele instalará `python-docx`.
4. Clique em **Pasta do Drive…** e escolha a pasta sincronizada em que quer guardar o caderno.

## Criar o executável
1. Com Python instalado, dê dois cliques em `Gerar executavel.bat`.
2. Aguarde a compilação.
3. O executável será criado em `dist\\Caderno Faculdade\\Caderno Faculdade.exe`. Copie a pasta inteira gerada para o local em que deseja usar o programa.

## Estrutura sugerida
```text
Google Drive/
└── Caderno Faculdade/
    ├── 1º Período/
    │   ├── Introdução ao Jornalismo/
    │   │   └── Ética e notícia.docx
    │   └── Teoria da Comunicação/
    │       └── Aula 01.docx
    └── 2º Período/
        └── ...
```

## Observações
- Para evitar sobrescrever arquivos por engano, o aplicativo pergunta antes de substituir um documento existente.
- Excluir uma matéria remove o arquivo `.docx` da pasta; se estiver sincronizada, essa exclusão também será sincronizada pelo Drive.
- Esta é uma primeira versão funcional. A edição preserva o texto dos parágrafos, mas não mantém formatação avançada, imagens, tabelas ou comentários de DOCX editados externamente.
