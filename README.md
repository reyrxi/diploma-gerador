# 📄 Gerador de Diploma e Histórico Escolar

Aplicativo desktop para preenchimento automático de modelos Word de **Diploma** e **Histórico Escolar**, desenvolvido em Python com interface gráfica.

---

## ✨ Funcionalidades

- Preenche automaticamente modelos `.docx` com os dados do aluno
- Suporte a **caixas de texto**, tabelas, cabeçalhos e rodapés
- Preserva a formatação original do documento (negrito, itálico, fontes)
- Suporte a até **40 disciplinas** com notas individuais
- Data por extenso gerada automaticamente (`01/03/2024` → `1 de março de 2024`)
- Dados do aluno compartilhados entre Diploma e Histórico — preencha uma vez, gere os dois
- Pode ser distribuído como `.exe` sem precisar instalar Python no PC de destino

---

## 🖥️ Interface

| Aba | Conteúdo |
|---|---|
| 👤 Dados do Aluno | Campos pessoais compartilhados entre os dois documentos |
| 📋 Histórico Escolar | Data e configurações específicas do histórico |
| 🎓 Diploma | Número de registro, livro, folha e dados específicos do diploma |
| 📚 Disciplinas | Notas/situação de cada disciplina (nome é só referência visual) |

---

## 📦 Requisitos

- Python 3.9 ou superior
- Windows 10/11

---

## 🚀 Instalação

**1. Clone o repositório**
```bash
git clone https://github.com/reyrxi/diploma-gerador

```

**2. Instale as dependências**
```bash
pip install python-docx pyinstaller
```

> Se `pip` não for reconhecido, use:
> ```bash
> py -m pip install python-docx pyinstaller
> ```

---

## ▶️ Como Rodar

```bash
python app.py
```

---

## 🔨 Como Gerar o `.EXE`

```bash
py -m PyInstaller --onefile --windowed --name GeradorDiploma app.py
```

O arquivo **`GeradorDiploma.exe`** será gerado em `dist\`.  
Esse arquivo pode ser copiado para qualquer computador Windows sem precisar instalar Python.

> ⚠️ O Windows pode exibir um aviso de segurança na primeira execução. Clique em **"Mais informações" → "Executar assim mesmo"**.

---

## 📄 Como Preparar os Modelos Word

Abra seu modelo `.docx` e substitua os campos variáveis pelos **placeholders** abaixo.  
Use **Localizar & Substituir** (`Ctrl+H`) no Word para fazer as trocas rapidamente.

O programa substitui os placeholders em **qualquer lugar do documento**: parágrafos normais, caixas de texto, tabelas, cabeçalhos e rodapés.

---

### 👤 Placeholders — Dados do Aluno (comuns aos dois documentos)

| Campo | Placeholder |
|---|---|
| Nome do(a) aluno(a) | `{{ALUNO}}` |
| Data de nascimento | `{{DATA_NASC}}` |
| Nacionalidade | `{{NACIONALIDADE}}` |
| Naturalidade (cidade) | `{{NATURALIDADE}}` |
| UF | `{{UF}}` |
| Naturalidade + UF combinados | `{{NATURALIDADE_UF}}` |
| Filiação 1 (1º responsável) | `{{FILIACAO_1}}` |
| Filiação 2 (2º responsável) | `{{FILIACAO_2}}` |
| CPF | `{{CPF}}` |
| RG | `{{RG}}` |
| Órgão emissor do RG | `{{ORGAO_EMISSOR}}` |
| Curso anterior | `{{CURSO_ANT}}` |
| Estabelecimento anterior | `{{ESTAB_ANT}}` |
| Ano de conclusão anterior | `{{ANO_ANT}}` |
| Cidade do curso anterior | `{{CIDADE_ANT}}` |
| Turma | `{{TURMA}}` |
| Data de início | `{{DATA_INICIO}}` |
| Data de término | `{{DATA_TERMINO}}` |
| Frequência (%) | `{{FREQUENCIA}}` |
| Resultado | `{{RESULTADO}}` |
| Código SISTEC | `{{COD_SISTEC}}` |
| Código Censo | `{{COD_CENSO}}` |
| Carga horária de estágio | `{{CARGA_ESTAGIO}}` |

---

### 📋 Placeholders — Histórico Escolar

| Campo | Placeholder |
|---|---|
| Data por extenso | `{{DATA_HIST}}` |
| Data no formato curto | `{{DATA_HIST_CURTA}}` |

**Notas das disciplinas** — as disciplinas já estão no modelo, coloque apenas o placeholder da nota na célula correspondente:

| Disciplina | Placeholder da nota |
|---|---|
| 1ª disciplina | `{{NOTA_1}}` |
| 2ª disciplina | `{{NOTA_2}}` |
| 3ª disciplina | `{{NOTA_3}}` |
| *(até 40)* | `{{NOTA_40}}` |

---

### 🎓 Placeholders — Diploma

| Campo | Placeholder |
|---|---|
| Data do diploma por extenso | `{{DATA_DIPLOMA}}` |
| Data do diploma (curta) | `{{DATA_DIPLOMA_CURTA}}` |
| Data de conclusão do curso | `{{DATA_CONCLUSAO}}` |
| Data de expedição do RG | `{{EXPEDIDO_EM}}` |
| Número do registro | `{{NUM_REGISTRO}}` |
| Folha | `{{FOLHA}}` |
| Livro nº | `{{LIVRO}}` |
| Nome da escola emissora | `{{NOME_ESCOLA}}` |
| Município/UF da escola | `{{MUNICIPIO_UF}}` |

---

## 💡 Dicas

- O campo **"Disciplina"** na aba Disciplinas é apenas referência visual — o nome não é inserido no documento, somente a nota.
- Use **"Gerar Ambos"** para gerar Histórico e Diploma em sequência sem precisar preencher os dados duas vezes.
- Certifique-se de que o modelo está no formato `.docx` (não `.doc`). Para converter, abra o arquivo no Word e salve como `.docx`.

---

## ❓ Problemas Comuns

| Problema | Solução |
|---|---|
| `pip` não reconhecido | Use `py -m pip install ...` ou reinstale o Python marcando **"Add to PATH"** |
| Placeholder não substituído | Verifique se está escrito exatamente igual, incluindo `{{` e `}}` |
| Formatação sumiu | O placeholder no modelo deve ter a formatação desejada (ex: negrito) |
| Arquivo não abre | O modelo precisa ser `.docx`, não `.doc` |
| Antivírus bloqueia o `.exe` | Adicione uma exceção no antivírus — falso positivo comum com PyInstaller |
| Janela fecha sozinha ao abrir o `.exe` | Gere novamente sem `--windowed` para ver o erro no terminal |
| Erro ao gerar EXE | Execute `py -m pip install --upgrade pyinstaller` e tente novamente |

---

## 🛠️ Tecnologias

- [Python](https://python.org) — linguagem principal
- [python-docx](https://python-docx.readthedocs.io) — manipulação de arquivos Word
- [tkinter](https://docs.python.org/3/library/tkinter.html) — interface gráfica (incluso no Python)
- [PyInstaller](https://pyinstaller.org) — geração do executável `.exe`
