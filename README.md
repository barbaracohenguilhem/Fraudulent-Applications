# DataSet limpo em partes

Dataset limpo de **anúncios de vagas de emprego**, com a indicação de quais são fraudulentos. Os dados foram divididos em 9 arquivos CSV (`DataSet_parte_01.csv` a `DataSet_parte_09.csv`, ~5 MB cada) para facilitar o armazenamento e o versionamento no GitHub.

## Estrutura do repositório

| Arquivo | Descrição |
| --- | --- |
| [`app_dash.py`](./app_dash.py) | **Dashboard Interativo em Python Dash** conectado ao banco SQLite (`http://127.0.0.1:8050`) |
| [`Database_Excel_1_ao_5.xlsx`](./Database_Excel_1_ao_5.xlsx) | **Database em Excel** com AutoFiltro, fórmulas `SUBTOTAL` dinâmicas e aba `Resumo_Dashboard` com gráficos nativos |
| [`DataSet_1_ao_5.db`](./DataSet_1_ao_5.db) | **Banco de Dados SQLite** (`tabela vagas`, 10.000 registros limpos das partes 1 a 5) |
| [`DataSet_1_ao_5_limpo.csv`](./DataSet_1_ao_5_limpo.csv) | Base consolidada limpa (partes 1 a 5) sem ruído de HTML/CSS e com hashes padronizados (`[URL]`, `[EMAIL]`, `[TELEFONE]`) |
| [`limpar_1_ao_5.py`](./limpar_1_ao_5.py) | Script de limpeza profunda de texto, encoding e consolidação das partes 1 a 5 |
| [`gerar_databases.py`](./gerar_databases.py) | Script gerador da Database SQLite e da planilha interativa do Excel |
| [`Entrega_Professor.zip`](./Entrega_Professor.zip) | Pacote completo consolidado para entrega |
| `DataSet_parte_01.csv` … `DataSet_parte_09.csv` | Partes do dataset; as partes `01` a `05` já estão higienizadas |

## Dashboard Interativo (Python Dash) & Database Excel

### Links Rápidos

- **Acesso Local ao Dashboard Dash:** [http://127.0.0.1:8050](http://127.0.0.1:8050)
- **Código-fonte do Dashboard:** [`app_dash.py`](./app_dash.py)
- **Planilha Database Excel Interativa:** [`Database_Excel_1_ao_5.xlsx`](./Database_Excel_1_ao_5.xlsx)
- **Banco de Dados SQLite:** [`DataSet_1_ao_5.db`](./DataSet_1_ao_5.db)

### Como executar o Dashboard Python Dash

```bash
pip install -r requirements.txt
python app_dash.py
```

Depois, abra **[http://127.0.0.1:8050](http://127.0.0.1:8050)** no navegador (ou dê dois cliques em [`ABRIR_DASHBOARD.command`](./ABRIR_DASHBOARD.command) no macOS).

### Funcionalidades do Dashboard (`app_dash.py`)

- **10 filtros interativos combinados:** Status (`Todas` / `Reais` / `Fraudulentas`), Busca textual por palavra-chave, Presença de Logo, Trabalho Remoto, País, Setor/Indústria, Tipo de Contrato, Experiência, Escolaridade e Faixa de Tamanho da Descrição.
- **5 KPIs dinâmicos + 4 gráficos Plotly interativos** em tempo real.
- **Tabela da Database com busca/ordenação por coluna** e **Inspetor de Vaga** (clique em qualquer linha para ler a descrição, requisitos e benefícios completos já limpos).
- **Exportação instantânea** do recorte filtrado para `.csv`.

## Colunas

| Coluna | Descrição |
| --- | --- |
| `title` | Título da vaga |
| `location` | Localização (país, estado, cidade) |
| `department` | Departamento |
| `salary_range` | Faixa salarial |
| `company_profile` | Perfil da empresa |
| `description` | Descrição da vaga |
| `requirements` | Requisitos |
| `benefits` | Benefícios |
| `telecommuting` | Indica se a vaga permite trabalho remoto (0/1) |
| `has_company_logo` | Indica se o anúncio tem logotipo da empresa (0/1) |
| `has_questions` | Indica se o anúncio tem perguntas de triagem (0/1) |
| `employment_type` | Tipo de contrato (ex.: tempo integral, meio período) |
| `required_experience` | Nível de experiência exigido |
| `required_education` | Nível de escolaridade exigido |
| `industry` | Setor da empresa |
| `function` | Função/área da vaga |
| `fraudulent` | **Variável alvo**: 1 se o anúncio é fraudulento, 0 caso contrário |
| `in_balanced_dataset` | Indica se o registro pertence ao subconjunto balanceado |

> Os campos de texto podem conter quebras de linha e vírgulas (entre aspas), e os arquivos começam com BOM UTF-8. Use um leitor de CSV adequado.

## Como usar

Carregue e junte todas as partes com pandas:

```python
import glob
import pandas as pd

arquivos = sorted(glob.glob("DataSet_parte_*.csv"))
df = pd.concat((pd.read_csv(a, encoding="utf-8-sig") for a in arquivos), ignore_index=True)

print(df.shape)
print(df["fraudulent"].value_counts())
```

## Possíveis usos

- Detecção de vagas de emprego fraudulentas
- Classificação de texto / NLP
- Análise exploratória de anúncios de emprego
