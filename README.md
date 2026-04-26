# Gerador de Diploma e Histórico Escolar
## Manual de Uso e Instalação

---

## 📦 Requisitos
- Python 3.9 ou superior  
- Windows 10/11 (para gerar o .EXE)

---

## 🚀 Como Instalar as Dependências

Abra o **Prompt de Comando** (cmd) na pasta do projeto e execute:

```
pip install python-docx pyinstaller
```

---

## ▶️ Como Rodar (sem compilar)

```
python app.py
```

---

## 🔨 Como Gerar o Arquivo .EXE

Execute no cmd:

```
pyinstaller diploma_app.spec
```

O arquivo **GeradorDiploma.exe** será criado na pasta `dist\`.  
Copie apenas esse arquivo para qualquer computador — não precisa de Python instalado!

> Se preferir um comando rápido sem o .spec:
> ```
> pyinstaller --onefile --windowed --name GeradorDiploma app.py
> ```

---

## 📄 Como Preparar os Modelos Word

Abra seus modelos `.docx` e substitua os textos variáveis pelos **placeholders** abaixo.
Eles serão substituídos automaticamente pelo programa.

> **Dica:** Use Localizar & Substituir (Ctrl+H) no Word para trocar os textos antigos.

### Histórico Escolar
| Campo no Word               | Placeholder             |
|-----------------------------|-------------------------|
| Nome do aluno               | `{{ALUNO}}`             |
| Data de nascimento          | `{{DATA_NASC}}`         |
| Nacionalidade               | `{{NACIONALIDADE}}`     |
| Naturalidade                | `{{NATURALIDADE}}`      |
| UF                          | `{{UF}}`                |
| Filiação                    | `{{FILIACAO}}`          |
| CPF                         | `{{CPF}}`               |
| RG                          | `{{RG}}`                |
| Órgão emissor               | `{{ORGAO_EMISSOR}}`     |
| Curso anterior              | `{{CURSO_ANT}}`         |
| Estabelecimento anterior    | `{{ESTAB_ANT}}`         |
| Ano de conclusão anterior   | `{{ANO_ANT}}`           |
| Cidade anterior             | `{{CIDADE_ANT}}`        |
| Turma                       | `{{TURMA}}`             |
| Data início                 | `{{DATA_INICIO}}`       |
| Data término                | `{{DATA_TERMINO}}`      |
| Frequência                  | `{{FREQUENCIA}}`        |
| Resultado                   | `{{RESULTADO}}`         |
| Código SISTEC               | `{{COD_SISTEC}}`        |
| Código Censo                | `{{COD_CENSO}}`         |
| Carga horária estágio       | `{{CARGA_ESTAGIO}}`     |
| Data por extenso            | `{{DATA_HIST}}`         |
| 1ª disciplina – nome        | `{{DISC_1_NOME}}`       |
| 1ª disciplina – média       | `{{DISC_1_NOTA}}`       |
| 2ª disciplina – nome        | `{{DISC_2_NOME}}`       |
| 2ª disciplina – média       | `{{DISC_2_NOTA}}`       |
| *(até 40 disciplinas)*      | `{{DISC_40_NOME}}` etc. |

### Diploma
| Campo no Word               | Placeholder              |
|-----------------------------|--------------------------|
| Nome do aluno               | `{{ALUNO}}`              |
| Naturalidade/UF             | `{{NATURALIDADE_UF}}`    |
| Data de nascimento          | `{{DATA_NASC}}`          |
| RG                          | `{{RG}}`                 |
| Órgão emissor               | `{{ORGAO_EMISSOR}}`      |
| Expedido em                 | `{{EXPEDIDO_EM}}`        |
| CPF                         | `{{CPF}}`                |
| Curso anterior              | `{{CURSO_ANT}}`          |
| Ano curso anterior          | `{{ANO_ANT}}`            |
| Estabelecimento             | `{{ESTAB_ANT}}`          |
| Nome da escola emissora     | `{{NOME_ESCOLA}}`        |
| Município/UF escola         | `{{MUNICIPIO_UF}}`       |
| Número do registro          | `{{NUM_REGISTRO}}`       |
| Folha                       | `{{FOLHA}}`              |
| Livro nº                    | `{{LIVRO}}`              |
| Data de conclusão           | `{{DATA_CONCLUSAO}}`     |
| Data do diploma (extenso)   | `{{DATA_DIPLOMA}}`       |
| Data do diploma (curta)     | `{{DATA_DIPLOMA_CURTA}}` |
| Código SISTEC               | `{{COD_SISTEC}}`         |
| Código Censo                | `{{COD_CENSO}}`          |
| Carga horária estágio       | `{{CARGA_ESTAGIO}}`      |

---

## 💡 Dicas

- **Dados do Aluno** são compartilhados entre Diploma e Histórico — preencha uma vez, gere os dois.
- Use **"Gerar Ambos"** para selecionar os dois modelos em sequência e gerar os dois documentos.
- O app suporta até **40 disciplinas** no Histórico.
- A data por extenso é gerada automaticamente: `01/03/2024` → `1 de março de 2024`.

---

## ❓ Problemas Comuns

| Problema | Solução |
|----------|---------|
| Placeholder não foi substituído | Verifique se o texto está exatamente igual, incluindo `{{` e `}}` |
| Arquivo não abre | Certifique-se de que o modelo é `.docx` (não `.doc`) |
| Erro ao gerar EXE | Rode `pip install --upgrade pyinstaller` |
