# DataSet limpo em partes

Dataset limpo de **anúncios de vagas de emprego**, com a indicação de quais são fraudulentos. Os dados foram divididos em 9 arquivos CSV (`DataSet_parte_01.csv` a `DataSet_parte_09.csv`, ~5 MB cada) para facilitar o armazenamento e o versionamento no GitHub.

## Estrutura do repositório

| Arquivo | Descrição |
|---|---|
| `DataSet_parte_01.csv` … `DataSet_parte_09.csv` | Partes do dataset; todas têm o mesmo cabeçalho |

## Colunas

| Coluna | Descrição |
|---|---|
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
