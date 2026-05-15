import tkinter as tk
from tkinter import messagebox, ttk

from constants import DIPLOMA_PLACEHOLDERS, HISTORICO_PLACEHOLDERS
from ui.fields import make_field, make_label_section
from ui.widgets import ScrollFrame


def build_tab_geral(notebook) -> dict:
    frame = ttk.Frame(notebook)
    notebook.add(frame, text="\U0001f464 Dados do Aluno")
    sf = ScrollFrame(frame, bg="#f5f6fa")
    sf.pack(fill="both", expand=True)
    f = sf.inner
    f.columnconfigure(1, weight=1)
    vars = {}
    def sv(key):
        vars[key] = tk.StringVar()
        return vars[key]
    r = 0
    r = make_label_section(f, r, "Identificacao do(a) Aluno(a)")
    make_field(f, r, "Aluno(a):",              sv("aluno"),         tooltip="Nome completo do aluno"); r += 1
    make_field(f, r, "Data de Nascimento:",    sv("data_nasc"),     tooltip="Formato: DD/MM/AAAA"); r += 1
    make_field(f, r, "Nacionalidade:",         sv("nacionalidade")); r += 1
    make_field(f, r, "Naturalidade:",          sv("naturalidade")); r += 1
    make_field(f, r, "UF:",                    sv("uf"),            width=8); r += 1
    make_field(f, r, "Filiacao 1 (Pai/Mae):",  sv("filiacao_1"),    width=55); r += 1
    make_field(f, r, "Filiacao 2 (Pai/Mae):",  sv("filiacao_2"),    width=55); r += 1
    make_field(f, r, "CPF:",                   sv("cpf"),           tooltip="Formato: 000.000.000-00"); r += 1
    make_field(f, r, "RG:",                    sv("rg")); r += 1
    make_field(f, r, "Orgao Emissor:",         sv("orgao_emissor")); r += 1
    r = make_label_section(f, r, "Curso Anterior")
    make_field(f, r, "Curso:",                 sv("curso_ant")); r += 1
    make_field(f, r, "Estabelecimento:",       sv("estab_ant"),     width=55); r += 1
    make_field(f, r, "Ano:",                   sv("ano_ant"),       width=10); r += 1
    make_field(f, r, "Cidade:",                sv("cidade_ant")); r += 1
    r = make_label_section(f, r, "Dados do Curso Atual")
    make_field(f, r, "Turma:",                 sv("turma"),         width=20); r += 1
    make_field(f, r, "Data de Inicio:",        sv("data_inicio"),   tooltip="Formato: DD/MM/AAAA"); r += 1
    make_field(f, r, "Data de Termino:",       sv("data_termino"),  tooltip="Formato: DD/MM/AAAA"); r += 1
    make_field(f, r, "Frequencia (%):",        sv("frequencia"),    width=10); r += 1
    make_field(f, r, "Resultado:",             sv("resultado")); r += 1
    make_field(f, r, "Codigo SISTEC:",         sv("cod_sistec")); r += 1
    make_field(f, r, "Codigo Censo:",          sv("cod_censo")); r += 1
    make_field(f, r, "Carga Horaria Estagio:", sv("carga_estagio"), width=55); r += 1
    return vars


def build_tab_historico(notebook) -> dict:
    frame = ttk.Frame(notebook)
    notebook.add(frame, text="\U0001f4cb Historico Escolar")
    sf = ScrollFrame(frame, bg="#f5f6fa")
    sf.pack(fill="both", expand=True)
    f = sf.inner
    f.columnconfigure(1, weight=1)
    vars = {}
    def sv(key):
        vars[key] = tk.StringVar()
        return vars[key]
    r = 0
    r = make_label_section(f, r, "Dados do Registro")
    make_field(f, r, "Data do Historico:", sv("data_hist"),
               tooltip="Formato: DD/MM/AAAA - sera convertida para extenso"); r += 1
    ttk.Label(f, text=(
        "Info: Todos os campos pessoais sao preenchidos na aba 'Dados do Aluno'.\n"
        "      Os campos de disciplinas sao preenchidos na aba 'Disciplinas'.\n"
        "      Configure os placeholders no seu modelo Word conforme o guia abaixo."
    ), font=("Segoe UI", 9), foreground="#555", background="#f5f6fa",
        justify="left", wraplength=580).grid(
        row=r, column=0, columnspan=2, padx=10, pady=6, sticky="w"); r += 1
    _build_placeholder_guide(f, r, HISTORICO_PLACEHOLDERS)
    return vars


def build_tab_diploma(notebook) -> dict:
    frame = ttk.Frame(notebook)
    notebook.add(frame, text="\U0001f393 Diploma")
    sf = ScrollFrame(frame, bg="#f5f6fa")
    sf.pack(fill="both", expand=True)
    f = sf.inner
    f.columnconfigure(1, weight=1)
    vars = {}
    def sv(key):
        vars[key] = tk.StringVar()
        return vars[key]
    r = 0
    r = make_label_section(f, r, "Dados do Registro do Diploma")
    make_field(f, r, "Numero do Registro:",   sv("num_registro"),  width=15); r += 1
    make_field(f, r, "Folha:",                sv("folha"),         width=10); r += 1
    make_field(f, r, "Livro nr:",             sv("livro"),         width=10); r += 1
    make_field(f, r, "Data do Diploma:",      sv("data_diploma"),
               tooltip="Formato: DD/MM/AAAA - sera convertida para extenso"); r += 1
    make_field(f, r, "Data de Conclusao:",    sv("data_conclusao"), tooltip="Formato: DD/MM/AAAA"); r += 1
    make_field(f, r, "Nome da Escola:",       sv("nome_escola"),   width=55); r += 1
    make_field(f, r, "Municipio/UF Escola:",  sv("municipio_uf"),  width=30); r += 1
    make_field(f, r, "Expedido em:",          sv("expedido_em"),
               tooltip="Data de expedicao do RG (DD/MM/AAAA)"); r += 1
    ttk.Label(f, text="Info: Os demais campos (nome, CPF, RG, etc.) vem da aba 'Dados do Aluno'.",
              font=("Segoe UI", 9), foreground="#555", background="#f5f6fa").grid(
        row=r, column=0, columnspan=2, padx=10, pady=6, sticky="w"); r += 1
    _build_placeholder_guide(f, r, DIPLOMA_PLACEHOLDERS)
    return vars


def _build_placeholder_guide(parent, row, placeholders):
    frame = ttk.LabelFrame(parent, text="Guia de Placeholders para o Modelo Word", padding=10)
    frame.grid(row=row, column=0, columnspan=2, sticky="ew", padx=8, pady=8)
    frame.columnconfigure(0, weight=1)
    text_box = tk.Text(frame, height=14, width=72, font=("Courier New", 9),
                       bg="#f0f4f8", relief="flat", wrap="none")
    text_box.pack(fill="both", expand=True)
    scrollbar = ttk.Scrollbar(frame, command=text_box.yview)
    scrollbar.pack(side="right", fill="y")
    text_box.configure(yscrollcommand=scrollbar.set)
    text_box.insert("end", placeholders)
    text_box.configure(state="disabled")
    def copy_to_clipboard():
        text_box.clipboard_clear()
        text_box.clipboard_append(placeholders)
        messagebox.showinfo("Copiado", "Texto copiado! Cole no seu editor de referencia.")
    ttk.Button(frame, text="Copiar", command=copy_to_clipboard).pack(anchor="e", pady=4)


class DisciplinasTab:
    """Aba de disciplinas com refresh dinamico da quantidade de linhas."""

    def __init__(self, notebook):
        self.vars = {}
        self.num_var = tk.IntVar(value=10)
        frame = ttk.Frame(notebook)
        notebook.add(frame, text="\U0001f4da Disciplinas")
        self._build(frame)

    def _build(self, parent):
        top = tk.Frame(parent, bg="#f5f6fa", pady=6)
        top.pack(fill="x", padx=10)
        ttk.Label(top, text="Quantidade de disciplinas:").pack(side="left")
        ttk.Spinbox(top, from_=1, to=40, textvariable=self.num_var,
                    width=4, command=self.refresh).pack(side="left", padx=6)
        ttk.Button(top, text="Atualizar", command=self.refresh).pack(side="left")
        tk.Label(
            parent,
            text="Info: O nome da disciplina eh apenas referencia visual. So a Nota/Situacao e inserida no documento.",
            font=("Segoe UI", 9), fg="#555", bg="#f5f6fa", anchor="w", justify="left",
        ).pack(fill="x", padx=12, pady=(0, 4))
        scroll = ScrollFrame(parent, bg="#f5f6fa")
        scroll.pack(fill="both", expand=True, padx=10, pady=4)
        self._inner = scroll.inner
        self._inner.columnconfigure(2, weight=1)
        self.refresh()

    def refresh(self):
        for widget in self._inner.winfo_children():
            widget.destroy()
        self.vars.clear()
        ttk.Label(self._inner, text="#", font=("Segoe UI", 9, "bold"), width=3).grid(
            row=0, column=0, padx=4, pady=2)
        ttk.Label(self._inner, text="Disciplina (referencia)",
                  font=("Segoe UI", 9, "bold")).grid(row=0, column=1, sticky="w", padx=4)
        ttk.Label(self._inner, text="Nota / Situacao -> inserida no doc",
                  font=("Segoe UI", 9, "bold"), foreground="#154360").grid(
            row=0, column=2, sticky="w", padx=4)
        for i in range(self.num_var.get()):
            label_var = tk.StringVar()
            nota_var = tk.StringVar()
            self.vars[f"nota_{i + 1}"] = nota_var
            self.vars[f"_label_{i + 1}"] = label_var
            ttk.Label(self._inner, text=f"{i + 1}", width=3, anchor="e").grid(
                row=i + 1, column=0, sticky="e", padx=(4, 2), pady=2)
            ttk.Entry(self._inner, textvariable=label_var, width=34,
                      foreground="#777").grid(row=i + 1, column=1, sticky="w", padx=4, pady=2)
            ttk.Entry(self._inner, textvariable=nota_var, width=18,
                      font=("Segoe UI", 10, "bold")).grid(
                row=i + 1, column=2, sticky="w", padx=4, pady=2)

    def get_nota_vars(self) -> dict:
        return {k: v for k, v in self.vars.items() if not k.startswith("_label_")}

    def clear(self):
        for var in self.vars.values():
            var.set("")
