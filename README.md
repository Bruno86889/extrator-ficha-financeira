# Extrator Financeiro

Aplicação dektop desenvolvida para automatizar a extração, organização e consolidação de informações de carga horária de professores do Centro Paula Souza (CPS), a partir de fichas financeiras exportadas da Prodesp em formato PDF.

O projeto tem como objetivo reduzir o trabalho manual de consulta de demonstrativos mensais, permitindo extrair todos os Vencimentos/Descontos (VDs) e quantidades de carga horária dos professores na folha de pagamento, gerando uma planilha com a carga por mês.

## Funcionalidades

- **Leitura de arquivos PDF**: processa folhas de pagamento para identificação dos dados
- Extração por VDs: extrai por padrão os Vds que possuem carga horária da folha de pagamento, por padrão `HORA AULA SEMANAL - ETEC`, `HAE COORDENACAO SEMANAL`, `HAEO - SEMANAL` e `HAE OUTROS SEMANAL`
- **Geração de Relatório Excel**: exporta todos os lançamentos identificados por mês

- **Interface gráfica:** seleção de arquivos de entrada, configuração da extração e escolha do local de salvamento.

## Tecnologias utilizadas

| Tecnologia | Finalidade |
|---|---|
| Python 3.12+ | Linguagem principal |
| Tkinter | Interface gráfica desktop |
| openpyxl | Criação e formatação de planilhas Excel |
| pdfplumber | Extração de texto de arquivos PDF, caso utilizada na versão atual |

## Estrutura do projeto

```text
extrator-financeiro/
├── src/
│   ├── main.py
│   ├── constants.py
│   ├── models/
│   │   ├── dates.py
│   │   ├── models.py
│   │   ├── numbers.py
│   │   ├── text.py
│   ├── infrastructure/
│   │   ├── pdf_reader.py
│   │   └── excel_exporter.py
│   └── ui/
│           ├── app.py
├── .gitignore
├── requirements.txt
└── README.md
```

A estrutura é uma proposta de organização, não necessariamente a estrutura já existente no repositório.

### Responsabilidades dos módulos

- `main.py`: inicialização da aplicação.
- `constants.py`: configurações gerais, rubricas e parâmetros de extração.
- `domain/`: utilitários e lógica da aplicação.
- `services/pdf_extractor.py`: leitura e extração de informações dos PDFs.
- `services/excel_generator.py`: geração e formatação do relatório Excel.
- `ui/`: componentes responsáveis pela interface gráfica.


## Requisitos de desenvolvimento

Antes de executar a aplicação, verifique se o ambiente possui:

- Python 3.15+.
- Pip para instalação das dependências.
- Arquivos PDF compatíveis com o processo de extração.
- Microsoft Excel ou outro aplicativo compatível com arquivos `.xlsx`, caso seja necessário visualizar os resultados.

## Instalação

### 1. Clone o repositório

```bash
git clone https://github.com/Bruno86889/extrator-ficha-financeira.git
cd extrator-financeiro
```

### 2. Crie um ambiente virtual

Windows:

```powershell
py -m venv .venv
```

Ative o ambiente:

```powershell
.venv\Scripts\Activate.ps1
```

Se estiver utilizando o Prompt de Comando:

```bat
.venv\Scripts\activate.bat
```

### 3. Instale as dependências

```bash
python -m pip install -r requirements.txt
```

## Como utilizar

1. Inicie a aplicação.
2. Selecione o arquivo PDF que contém a ficha financeira exportada da Prodesp.
3. Indique o nome do funcionário para aparecer no relatório
3. Verifique os VDs selecionados por padrão e altere para outros vds que deseja extrair, por exemplo, `HORA AULA SEMANAL - ETEC`.
4. Clique em exportar.
6. Selecione a pasta de destino.
7. Gere a planilha Excel com os lançamentos individuais e o resumo mensal.


## Funcionalidades Planejadas
- [ ] Importação de arquivos em lote, com relatório consolidado de todos os funcionários.
- [ ] Extração de faltas por funcionário.
- [ ] Melhoria na interface do usuário.
