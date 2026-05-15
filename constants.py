HISTORICO_PLACEHOLDERS = """\
PLACEHOLDERS PARA O MODELO DO HISTÓRICO ESCOLAR
Copie e cole estes marcadores exatamente no seu documento Word
onde o valor variável deve aparecer:

=== DADOS DO ALUNO ===
{{ALUNO}}             → Nome completo do(a) aluno(a)
{{DATA_NASC}}         → Data de nascimento (DD/MM/AAAA)
{{NACIONALIDADE}}     → Nacionalidade
{{NATURALIDADE}}      → Naturalidade (cidade)
{{UF}}                → UF de naturalidade
{{FILIACAO_1}}        → Filiação 1 (1º responsável)
{{FILIACAO_2}}        → Filiação 2 (2º responsável)
{{CPF}}               → CPF
{{RG}}                → RG
{{ORGAO_EMISSOR}}     → Órgão emissor do RG

=== CURSO ANTERIOR ===
{{CURSO_ANT}}         → Curso anterior
{{ESTAB_ANT}}         → Estabelecimento
{{ANO_ANT}}           → Ano de conclusão
{{CIDADE_ANT}}        → Cidade

=== DADOS DO CURSO ===
{{TURMA}}             → Turma
{{DATA_INICIO}}       → Data de início
{{DATA_TERMINO}}      → Data de término
{{FREQUENCIA}}        → Frequência (%)
{{RESULTADO}}         → Resultado
{{COD_SISTEC}}        → Código SISTEC
{{COD_CENSO}}         → Código Censo
{{CARGA_ESTAGIO}}     → Carga horária de estágio

=== DATA ===
{{DATA_HIST}}         → Ex.: 1 de março de 2024 (por extenso)
{{DATA_HIST_CURTA}}   → Ex.: 01/03/2024

=== NOTAS DAS DISCIPLINAS ===
As disciplinas já estão no modelo — coloque apenas o
placeholder da nota na célula/campo correspondente:

{{NOTA_1}}            → Nota/Situação da 1ª disciplina
{{NOTA_2}}            → Nota/Situação da 2ª disciplina
{{NOTA_3}}            → Nota/Situação da 3ª disciplina
... (até {{NOTA_40}})
"""

DIPLOMA_PLACEHOLDERS = """\
PLACEHOLDERS PARA O MODELO DO DIPLOMA
Copie e cole estes marcadores exatamente no seu documento Word:

=== DADOS DO ALUNO ===
{{ALUNO}}             → Nome completo
{{NATURALIDADE_UF}}   → Cidade/UF  (ex.: Rio de Janeiro/RJ)
{{NATURALIDADE}}      → Apenas a cidade
{{UF}}                → Apenas a UF
{{DATA_NASC}}         → Data de nascimento
{{RG}}                → Número do RG
{{ORGAO_EMISSOR}}     → Órgão emissor do RG
{{EXPEDIDO_EM}}       → Data de expedição do RG
{{CPF}}               → CPF

=== CURSO ANTERIOR ===
{{CURSO_ANT}}         → Ex.: ENSINO MÉDIO
{{ANO_ANT}}           → Ano de conclusão
{{ESTAB_ANT}}         → Estabelecimento

=== REGISTRO ===
{{NUM_REGISTRO}}      → Número do registro
{{FOLHA}}             → Folha
{{LIVRO}}             → Livro nº
{{NOME_ESCOLA}}       → Nome da escola emissora
{{MUNICIPIO_UF}}      → Município e UF da escola

=== DATAS ===
{{DATA_CONCLUSAO}}    → Data de conclusão do curso (DD/MM/AAAA)
{{DATA_DIPLOMA}}      → Ex.: 01 de março de 2024 (por extenso)
{{DATA_DIPLOMA_CURTA}} → Ex.: 01/03/2024

=== OUTROS ===
{{COD_SISTEC}}        → Código SISTEC
{{COD_CENSO}}         → Código Censo
{{CARGA_ESTAGIO}}     → Carga horária de estágio
"""
