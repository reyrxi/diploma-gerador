# Gerador de Diploma e Historico Escolar

Aplicativo desktop para preenchimento automatico de modelos Word de **Diploma** e **Historico Escolar**, desenvolvido em Python com interface grafica.

---

## Funcionalidades

- Preenche automaticamente modelos `.docx` com os dados do aluno
- Suporte a **caixas de texto**, tabelas, cabecalhos e rodapes
- Preserva a formatacao original do documento (negrito, italico, fontes)
- Suporte a ate **40 disciplinas** com notas individuais
- Data por extenso gerada automaticamente (`01/03/2024` -> `1 de marco de 2024`)
- Dados do aluno compartilhados entre Diploma e Historico - preencha uma vez, gere os dois
- Pode ser distribuido como `.exe` sem precisar instalar Python no PC de destino

---

## Interface

| Aba | Conteudo |
|---|---|
| Dados do Aluno | Campos pessoais compartilhados entre os dois documentos |
| Historico Escolar | Data e configuracoes especificas do historico |
| Diploma | Numero de registro, livro, folha e dados especificos do diploma |
| Disciplinas | Notas/situacao de cada disciplina (nome e so referencia visual) |

---

## Estrutura do Projeto

```
diploma-gerador/
+-- main.py              # Ponto de entrada
+-- constants.py         # Textos dos guias de placeholders
+-- requirements.txt     # Dependencias Python
+-- core/
|   +-- document.py      # Substituicao de placeholders no .docx
|   +-- dates.py         # Formatacao de datas por extenso
+-- ui/
    +-- app.py           # Janela principal (DiplomaApp)
    +-- tabs.py          # Construtores de abas e DisciplinasTab
    +-- widgets.py       # ToolTip e ScrollFrame
    +-- fields.py        # Helpers de campos de formulario
```

---

## Requisitos

- Python 3.10 ou superior
- Windows 10/11

---

## Instalacao

**1. Clone o repositorio**
```bash
git clone https://github.com/reyrxi/diploma-gerador
cd diploma-gerador
```

**2. Instale as dependencias**
```bash
pip install -r requirements.txt
```

> Se `pip` nao for reconhecido, use:
> ```bash
> py -m pip install -r requirements.txt
> ```

---

## Como Rodar

```bash
python main.py
```

---

## Como Gerar o `.EXE`

```bash
py -m PyInstaller --onefile --windowed --name GeradorDiploma main.py
```

O arquivo **`GeradorDiploma.exe`** sera gerado em `dist\`.
Esse arquivo pode ser copiado para qualquer computador Windows sem precisar instalar Python.

> O Windows pode exibir um aviso de seguranca na primeira execucao. Clique em **"Mais informacoes" -> "Executar assim mesmo"**.

---

## Como Preparar os Modelos Word

Abra seu modelo `.docx` e substitua os campos variaveis pelos **placeholders** abaixo.
Use **Localizar & Substituir** (`Ctrl+H`) no Word para fazer as trocas rapidamente.

O programa substitui os placeholders em **qualquer lugar do documento**: paragrafos normais, caixas de texto, tabelas, cabecalhos e rodapes.

---

### Placeholders - Dados do Aluno (comuns aos dois documentos)

| Campo | Placeholder |
|---|---|
| Nome do(a) aluno(a) | `{{ALUNO}}` |
| Data de nascimento | `{{DATA_NASC}}` |
| Nacionalidade | `{{NACIONALIDADE}}` |
| Naturalidade (cidade) | `{{NATURALIDADE}}` |
| UF | `{{UF}}` |
| Naturalidade + UF combinados | `{{NATURALIDADE_UF}}` |
| Filiacao 1 (1o responsavel) | `{{FILIACAO_1}}` |
| Filiacao 2 (2o responsavel) | `{{FILIACAO_2}}` |
| CPF | `{{CPF}}` |
| RG | `{{RG}}` |
| Orgao emissor do RG | `{{ORGAO_EMISSOR}}` |
| Curso anterior | `{{CURSO_ANT}}` |
| Estabelecimento anterior | `{{ESTAB_ANT}}` |
| Ano de conclusao anterior | `{{ANO_ANT}}` |
| Cidade do curso anterior | `{{CIDADE_ANT}}` |
| Turma | `{{TURMA}}` |
| Data de inicio | `{{DATA_INICIO}}` |
| Data de termino | `{{DATA_TERMINO}}` |
| Frequencia (%) | `{{FREQUENCIA}}` |
| Resultado | `{{RESULTADO}}` |
| Codigo SISTEC | `{{COD_SISTEC}}` |
| Codigo Censo | `{{COD_CENSO}}` |
| Carga horaria de estagio | `{{CARGA_ESTAGIO}}` |

---

### Placeholders - Historico Escolar

| Campo | Placeholder |
|---|---|
| Data por extenso | `{{DATA_HIST}}` |
| Data no formato curto | `{{DATA_HIST_CURTA}}` |

**Notas das disciplinas** - as disciplinas ja estao no modelo, coloque apenas o placeholder da nota na celula correspondente:

| Disciplina | Placeholder da nota |
|---|---|
| 1a disciplina | `{{NOTA_1}}` |
| 2a disciplina | `{{NOTA_2}}` |
| 3a disciplina | `{{NOTA_3}}` |
| *(ate 40)* | `{{NOTA_40}}` |

---

### Placeholders - Diploma

| Campo | Placeholder |
|---|---|
| Data do diploma por extenso | `{{DATA_DIPLOMA}}` |
| Data do diploma (curta) | `{{DATA_DIPLOMA_CURTA}}` |
| Data de conclusao do curso | `{{DATA_CONCLUSAO}}` |
| Data de expedicao do RG | `{{EXPEDIDO_EM}}` |
| Numero do registro | `{{NUM_REGISTRO}}` |
| Folha | `{{FOLHA}}` |
| Livro nr | `{{LIVRO}}` |
| Nome da escola emissora | `{{NOME_ESCOLA}}` |
| Municipio/UF da escola | `{{MUNICIPIO_UF}}` |

---

## Dicas

- O campo **"Disciplina"** na aba Disciplinas e apenas referencia visual - o nome nao e inserido no documento, somente a nota.
- Use **"Gerar Ambos"** para gerar Historico e Diploma em sequencia sem precisar preencher os dados duas vezes.
- Certifique-se de que o modelo esta no formato `.docx` (nao `.doc`). Para converter, abra o arquivo no Word e salve como `.docx`.

---

## Problemas Comuns

| Problema | Solucao |
|---|---|
| `pip` nao reconhecido | Use `py -m pip install -r requirements.txt` ou reinstale o Python marcando **"Add to PATH"** |
| Placeholder nao substituido | Verifique se esta escrito exatamente igual, incluindo `{{` e `}}` |
| Formatacao sumiu | O placeholder no modelo deve ter a formatacao desejada (ex: negrito) |
| Arquivo nao abre | O modelo precisa ser `.docx`, nao `.doc` |
| Antivirus bloqueia o `.exe` | Adicione uma excecao no antivirus - falso positivo comum com PyInstaller |
| Janela fecha sozinha ao abrir o `.exe` | Gere novamente sem `--windowed` para ver o erro no terminal |
| Erro ao gerar EXE | Execute `py -m pip install --upgrade pyinstaller` e tente novamente |

---

## Tecnologias

- [Python](https://python.org) - linguagem principal
- [python-docx](https://python-docx.readthedocs.io) - manipulacao de arquivos Word
- [tkinter](https://docs.python.org/3/library/tkinter.html) - interface grafica (incluso no Python)
- [PyInstaller](https://pyinstaller.org) - geracao do executavel `.exe`
