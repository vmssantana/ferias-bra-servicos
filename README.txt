SISTEMA DE FÉRIAS - VERSÃO ONLINE COM GOOGLE SHEETS

Arquivos principais:
- app.py -> aplicação online em Flask
- templates/index.html -> tela única para o colaborador
- requirements.txt -> dependências do projeto
- render.yaml -> configuração pronta para o Render

Planilha esperada:
Aba 1: BASE_COLABORADORES
Colunas: MATRICULA | NOME | UNIDADE | MES_FERIAS

Aba 2: RESPOSTAS_FERIAS
Será criada automaticamente se não existir.

Variáveis obrigatórias no Render:
- SPREADSHEET_ID
- GOOGLE_SERVICE_ACCOUNT_JSON
- ADMIN_TOKEN

Importante:
- compartilhe a planilha com o e-mail da conta de serviço do Google
- coloque o JSON da conta de serviço inteiro na variável GOOGLE_SERVICE_ACCOUNT_JSON
