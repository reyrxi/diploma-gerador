MONTHS = [
    "", "janeiro", "fevereiro", "março", "abril", "maio", "junho",
    "julho", "agosto", "setembro", "outubro", "novembro", "dezembro",
]


def format_date_full(date_str):
    """Converte DD/MM/AAAA para 'D de mês de AAAA' (ex.: 1 de março de 2024)."""
    try:
        d, m, y = date_str.strip().split("/")
        return f"{int(d)} de {MONTHS[int(m)]} de {y}"
    except Exception:
        return date_str
