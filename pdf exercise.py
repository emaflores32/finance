# -*- coding: utf-8 -*-
"""
Created on Sat Jul 11 13:07:11 2026

@author: Ema Flores
"""

from __future__ import annotations

import calendar
import json
import os
import re
import subprocess
import sys
import unicodedata
import urllib.error
import urllib.request
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

try:
    import pdfplumber
except ModuleNotFoundError as exc:
    raise SystemExit(
        "Falta la librería pdfplumber. Instalala con:\n"
        "  pip install pdfplumber\n"
        "y volvé a ejecutar el script con ese mismo intérprete de Python."
    ) from exc

# Evita errores de encoding al imprimir acentos en la consola de Windows.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


directory = Path(r"C:\Users\emafl\OneDrive\Desktop\Personal\ONs")
PDF_INDEX = os.getenv("PDF_INDEX", "0").strip()
CAPITAL = Decimal("100")

MESES = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
    "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
    "septiembre": 9, "setiembre": 9, "octubre": 10,
    "noviembre": 11, "diciembre": 12,
}

DATE_RE = re.compile(
    r"(\d{1,2})\s+de\s+([a-záéíóúü]+)\s+de\s+(\d{4})",
    re.IGNORECASE)


# ---------------------------------------------------------------------------
# Lectura y extracción del PDF
# ---------------------------------------------------------------------------

def get_pdfs(folder: Path):
    return sorted(folder.glob("*.pdf"), key=lambda p: p.name.lower())


def open_pdfs(path: Path) -> str:
    with pdfplumber.open(path) as pdf:
        return "\n".join((page.extract_text() or "") for page in pdf.pages[:2])


def clean_text(text: str) -> str:
    # Algunos PDFs entregan los acentos descompuestos (o + tilde combinada),
    # lo que rompe las regex; NFC los deja en un único carácter.
    text = unicodedata.normalize("NFC", text)
    return (text.replace("\u00a0", " ")
                .replace("\u2009", " ")
                .replace("\u202f", " ")
                .replace("\r\n", "\n")
                .replace("\r", "\n"))


def extract_company(text: str):
    first_part = text[:5000]

    pattern = re.compile(
        r'(?P<company>[A-ZÁÉÍÓÚÜÑ0-9\s,&\-.]+?S\.A(?:\.U\.?)?)'
        r'.*?'
        r'CLASE\s+(?P<bond>[A-Z0-9IVXLCDM]+)',
        re.DOTALL | re.IGNORECASE)

    match = pattern.search(first_part)
    if not match:
        return None

    company = " ".join(match.group("company").split())
    company = re.sub(r"^AVISO DE RESULTADOS\s+", "", company, flags=re.IGNORECASE)
    return {"company": company, "bond": match.group("bond").upper()}


def split_bonds_sections(text: str):
    pattern = re.compile(
        r'(Obligaciones Negociables(?:\s+Simples)?\s+Clase\s+[A-Z0-9IVXLCDM]+:.*?)'
        r'(?=Obligaciones Negociables(?:\s+Simples)?\s+Clase\s+[A-Z0-9IVXLCDM]+:|\Z)',
        re.DOTALL | re.IGNORECASE)
    return pattern.findall(text)


def extract_fields(text: str):
    """Extrae los puntos numerados del aviso de resultados.

    Los avisos usan tanto "1)" como "1." como numeración; el punto sólo se
    acepta si no va seguido de otro dígito, para no confundirlo con montos
    como "70.000.000". También se ignoran los números de tres o más dígitos.
    """
    pattern = re.compile(
        r"(?m)^\s*(\d{1,2})\s*(?:\)|\.(?!\d))\s*(.*?)"
        r"(?=^\s*\d{1,2}\s*(?:\)|\.(?!\d))\s|\Z)",
        re.DOTALL)
    return {int(num): value.strip() for num, value in pattern.findall(text)}


def find_dates(text: str):
    dates = []
    for match in DATE_RE.finditer(text):
        day, month_name, year = match.groups()
        month = MESES.get(month_name.lower())
        if month is None:
            continue
        try:
            value = date(int(year), month, int(day))
        except ValueError:
            continue
        if value not in dates:
            dates.append(value)
    return dates


def find_field(conditions, *patterns):
    """Busca el primer campo numerado cuyo texto coincida con algún patrón.

    Los avisos no numeran siempre igual los puntos (la fecha de emisión puede
    ser el punto 4, 6 u 8), así que no conviene depender de índices fijos.
    Algunos PDFs devuelven los acentos como "�"; por eso las vocales
    acentuadas de los patrones también aceptan ese carácter.
    """
    for text in conditions.values():
        for pattern in patterns:
            if re.match(_accent_tolerant(pattern), text, re.IGNORECASE):
                return text
    return ""


def _accent_tolerant(pattern: str) -> str:
    """Permite que las letras acentuadas matcheen cualquier carácter.

    Se usa junto a patrones escritos sin clases de caracteres (por ejemplo
    ``emisi.n``) para tolerar tanto acentos correctos como el carácter de
    reemplazo que devuelven algunos PDFs.
    """
    return re.sub(r"[áéíóúüñÁÉÍÓÚÜÑ]", ".", pattern)


# ---------------------------------------------------------------------------
# Cálculo local del cash flow (sin LLM)
# ---------------------------------------------------------------------------

def parse_fixed_rate(text: str) -> Decimal:
    if re.search(r"tasa\s+de\s+inter[eé]s\s+variable", text, re.IGNORECASE):
        raise ValueError(
            "el bono paga tasa variable; el cálculo local solo soporta tasa fija")

    match = re.search(
        r"(?:tasa\s+de\s+inter[eé]s|tasa\s+aplicable|tasa)[^%\d]*"
        r"([\d.,]+)\s*%",
        text, re.IGNORECASE)
    if not match:
        # Algunos avisos sólo traen el porcentaje suelto ("Tasa: 6,00%").
        match = re.search(r"([\d]+(?:[.,]\d+)?)\s*%", text, re.IGNORECASE)
    if not match:
        raise ValueError("no se pudo leer la tasa de interés del PDF")

    raw = match.group(1)
    if "," in raw:
        raw = raw.replace(".", "").replace(",", ".")
    return Decimal(raw) / Decimal(100)


def parse_amortization(text: str, maturity):
    dates = find_dates(text)
    percentages = [
        Decimal(value.replace(",", "."))
        for value in re.findall(r"(\d+(?:[.,]\d+)?)\s*%", text)
    ]

    if dates and len(dates) == len(percentages):
        total = sum(percentages, Decimal(0))
        if abs(total - Decimal(100)) <= Decimal("0.01"):
            return list(zip(dates, percentages))

    # Si no hay una tabla de amortización, se asume pago íntegro al vencimiento.
    return [(maturity, Decimal(100))]


def infer_frequency(text: str, dates) -> int:
    low = text.lower()

    if "trimestral" in low:
        return 4
    if "semestral" in low:
        return 2
    if re.search(r"\banual(?:mente)?\b", low):
        return 1

    unique = sorted(set(dates))
    if len(unique) >= 3:
        gaps = [(b - a).days for a, b in zip(unique, unique[1:])]
        gaps = [gap for gap in gaps if gap > 0]
        if gaps:
            average = sum(gaps) / len(gaps)
            return max(1, round(365 / average))

    if len(unique) == 2:
        first, last = unique
        if first.day == last.day and (last.month - first.month) % 6 == 0:
            return 1 if (last.month - first.month) % 12 == 0 else 2

    return 2


def add_months(value: date, months: int) -> date:
    month_index = value.month - 1 + months
    year = value.year + month_index // 12
    month = month_index % 12 + 1
    day = min(value.day, calendar.monthrange(year, month)[1])
    return date(year, month, day)


def get_coupon_dates(text: str, issue: date, maturity: date):
    explicit = find_dates(text)
    frequency = infer_frequency(text, explicit)

    # Si el aviso lista todas las fechas de pago, usarlas tal cual.
    # El vencimiento siempre es una fecha de pago (a veces se menciona como
    # "en la Fecha de Vencimiento" sin repetir el día/mes/año).
    if len(explicit) >= 3:
        return sorted(set(explicit) | {maturity}), frequency

    # Si solo aparece la fecha de vencimiento, generar desde la emisión.
    if explicit and explicit[0] >= maturity:
        start = add_months(issue, 12 // frequency)
    else:
        start = explicit[0] if explicit else add_months(issue, 12 // frequency)

    if start <= issue:
        start = add_months(issue, 12 // frequency)

    result = []
    current = start
    step = 12 // frequency
    while current < maturity:
        result.append(current)
        current = add_months(current, step)
    result.append(maturity)
    return sorted(set(result)), frequency


def days_30_360(start: date, end: date) -> int:
    """Días según convención 30/360, habitual en bonos corporativos."""
    day_start = min(start.day, 30)
    day_end = min(end.day, 30)
    return (360 * (end.year - start.year)
            + 30 * (end.month - start.month)
            + (day_end - day_start))


def money(value: Decimal) -> str:
    return f"{value:,.4f}".replace(",", "X").replace(".", ",").replace("X", ".")


def format_percent(rate: Decimal) -> str:
    return f"{rate * 100:.2f}%".replace(".", ",")


def compute_cash_flow(conditions, capital: Decimal = CAPITAL):
    """Calcula el cash flow del bono a partir de los campos del aviso.

    Devuelve ``(rows, rate, frequency)`` donde cada fila es
    ``(fecha, desde, días, capital, tasa, interés, amortización, total, saldo)``.
    """
    issue_text = find_field(
        conditions,
        r"fecha\s+de\s+emisi.n\s+y\s+liquidaci.n",
        r"fecha\s+de\s+emisi.n")
    maturity_text = find_field(conditions, r"fecha\s+de\s+vencimiento")
    amortization_text = find_field(conditions, r"amortizaci.n")
    coupon_text = find_field(
        conditions,
        r"fechas?\s+de\s+pago\s+de\s+intereses",
        r"pago\s+de\s+intereses")
    rate_text = find_field(
        conditions, r"tasa\s+de\s+inter.s", r"tasa\s+aplicable", r"tasa")

    issue_dates = find_dates(issue_text)
    maturity_dates = find_dates(maturity_text)

    if not issue_dates:
        raise ValueError("no se pudo leer la Fecha de Emisión y Liquidación")
    issue = issue_dates[0]

    if maturity_dates:
        maturity = maturity_dates[0]
    else:
        fallback_dates = find_dates(coupon_text)
        if not fallback_dates:
            raise ValueError("no se pudo leer la Fecha de Vencimiento")
        maturity = max(fallback_dates)

    rate = parse_fixed_rate(rate_text)
    amortization = parse_amortization(amortization_text, maturity)
    coupons, frequency = get_coupon_dates(coupon_text, issue, maturity)

    amortization_by_date = {d: p for d, p in amortization}

    outstanding = capital
    total_amortization = Decimal(0)
    total_interest = Decimal(0)
    accrued_interest = Decimal(0)
    last_accrual = issue
    rows = []

    for current in sorted(set(coupons) | set(amortization_by_date)):
        period_start = last_accrual
        period_days = days_30_360(period_start, current)
        opening_balance = outstanding

        if period_days > 0:
            accrued_interest += (
                outstanding * rate * Decimal(period_days) / Decimal(360)
            )
            last_accrual = current

        interest_payment = Decimal(0)
        if current in coupons:
            interest_payment = accrued_interest.quantize(
                Decimal("0.000001"), rounding=ROUND_HALF_UP)
            total_interest += interest_payment
            accrued_interest = Decimal(0)

        amortization_amount = Decimal(0)
        if current in amortization_by_date:
            amortization_amount = (
                capital * amortization_by_date[current] / Decimal(100)
            ).quantize(Decimal("0.000001"), rounding=ROUND_HALF_UP)
            amortization_amount = min(amortization_amount, outstanding)
            outstanding -= amortization_amount
            total_amortization += amortization_amount

        rows.append((
            current,
            period_start,
            period_days,
            opening_balance,
            rate,
            interest_payment,
            amortization_amount,
            amortization_amount + interest_payment,
            outstanding,
        ))

    return rows, rate, frequency


def build_cash_flow(conditions, capital: Decimal = CAPITAL) -> str:
    rows, rate, frequency = compute_cash_flow(conditions, capital)
    total_interest = sum((row[5] for row in rows), Decimal(0))
    total_amortization = sum((row[6] for row in rows), Decimal(0))
    total = total_amortization + total_interest

    lines = []
    lines.append(f"Tasa fija: {format_percent(rate)} nominal anual")
    lines.append(f"Capital inicial: USD {money(capital)}")
    lines.append(f"Frecuencia de intereses: {frequency} pago(s) por año")
    lines.append("")
    lines.append("Detalle de intereses y cash flow (convención 30/360)")
    lines.append(
        f"{'Fecha':<12} {'Desde':<12} {'Días':>5} {'Capital':>12} "
        f"{'Tasa':>8} {'Interés':>12} {'Amort.':>12} "
        f"{'Total':>12} {'Saldo':>12}")
    lines.append("-" * 116)

    for (current, period_start, period_days, opening_balance, row_rate,
         interest, amortization_amount, cash, balance) in rows:
        lines.append(
            f"{current.isoformat():<12} "
            f"{period_start.isoformat():<12} "
            f"{period_days:>5} "
            f"{money(opening_balance):>12} "
            f"{format_percent(row_rate):>8} "
            f"{money(interest):>12} "
            f"{money(amortization_amount):>12} "
            f"{money(cash):>12} "
            f"{money(balance):>12}")

    lines.append("-" * 116)
    lines.append(
        f"{'TOTALES':<12} {'':<12} {'':>5} {'':>12} {'':>8} "
        f"{money(total_interest):>12} "
        f"{money(total_amortization):>12} "
        f"{money(total):>12}")
    lines.append("")
    lines.append(
        "Cálculo del interés: capital pendiente × tasa anual × días/360. "
        "El capital es el saldo antes de la amortización de esa fecha.")
    lines.append(
        "Supuestos: capital de USD 100, tasa nominal anual y convención 30/360.")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Backends LLM (opcionales)
# ---------------------------------------------------------------------------

# Backend de cálculo:
#   local    (default) -> calcula el cash flow con los datos del PDF, sin internet
#   opencode            -> modelos free de OpenCode (requiere cuota disponible)
#   ollama              -> modelo local (requiere Ollama corriendo)
#   openai              -> API compatible OpenAI (requiere LLM_API_KEY)
BACKEND = os.getenv("LLM_BACKEND", "local").strip().lower()

OPENCODE_EXE = (
    Path(os.environ.get("APPDATA", ""))
    / "npm" / "node_modules" / "@opencode" / "cli" / "bin" / "opencode.exe"
)


def _post_json(url, payload, headers=None):
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    request = urllib.request.Request(
        url, data=data,
        headers={"Content-Type": "application/json", **(headers or {})})
    with urllib.request.urlopen(request, timeout=300) as response:
        return json.loads(response.read().decode("utf-8"))


def _ask_opencode(prompt):
    model = os.getenv("OPENCODE_MODEL", "deepseek-v4.1-flash-free")

    if not OPENCODE_EXE.exists():
        raise RuntimeError(f"no se encontró el ejecutable de OpenCode: {OPENCODE_EXE}")

    body = json.dumps(
        {"prompt": prompt, "model": {"providerID": "opencode", "id": model}},
        ensure_ascii=False)

    run = subprocess.run(
        [str(OPENCODE_EXE), "api", "post", "/api/experimental/generate/text",
         "-d", body],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=300)

    if run.returncode != 0 and not run.stdout.strip():
        raise RuntimeError(run.stderr.strip() or f"OpenCode terminó con código {run.returncode}")

    try:
        result = json.loads(run.stdout or "{}")
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"respuesta inválida de OpenCode: {run.stdout[:500]!r}") from exc

    if "data" not in result:
        if result.get("reason") == "rate_limit":
            wait = int(result.get("retryAfterMs", 0)) // 60000
            raise RuntimeError(
                f"cuota free de OpenCode agotada; reintentá en ~{wait} min "
                f"o usá LLM_BACKEND=local / ollama / openai")
        raise RuntimeError(result.get("message") or result)

    return result["data"]["text"]


def _ask_ollama(prompt):
    try:
        result = _post_json("http://localhost:11434/api/chat", {
            "model": os.getenv("OLLAMA_MODEL", "qwen3:4b"),
            "messages": [{"role": "user", "content": prompt}],
            "stream": False})
    except urllib.error.URLError as exc:
        raise RuntimeError(
            "Ollama no está corriendo en http://localhost:11434. "
            "Instalalo/abrilo o usá LLM_BACKEND=local") from exc

    if "error" in result:
        raise RuntimeError(result["error"])
    return result["message"]["content"]


def _ask_openai(prompt):
    api_key = os.getenv("LLM_API_KEY", "").strip()
    if not api_key:
        raise RuntimeError("falta la variable LLM_API_KEY para el backend openai")

    base = os.getenv("LLM_BASE_URL", "https://api.groq.com/openai/v1")
    result = _post_json(base.rstrip("/") + "/chat/completions", {
        "model": os.getenv("LLM_MODEL", "llama-3.3-70b-versatile"),
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0},
        {"Authorization": f"Bearer {api_key}"})
    return result["choices"][0]["message"]["content"]


def ask_llm(prompt):
    if BACKEND == "opencode":
        return _ask_opencode(prompt)
    if BACKEND == "ollama":
        return _ask_ollama(prompt)
    if BACKEND == "openai":
        return _ask_openai(prompt)
    raise ValueError(f"LLM_BACKEND desconocido: {BACKEND}")


def build_prompt(title, sections, conditions):
    company = title["company"] if title else "(empresa no detectada)"
    bond = title["bond"] if title else "(clase no detectada)"
    context = "\n\n".join(sections[:3])[:6000] if sections else "(sin secciones adicionales)"

    return f"""
sos un analista financiero revisando bonos corporativos.

Empresa: {company}
Bono: Clase {bond}

Condiciones:
{json.dumps(conditions, indent=2, ensure_ascii=False)}

Fragmentos del prospecto:
{context}

arma el cash flow del bono indicando amortizaciones y tasa de interes mas la sumatoria
asumiendo un capital de 100 usd. Devolvé una tabla con fecha, amortización, interés,
total y saldo pendiente.
""".strip()


# ---------------------------------------------------------------------------
# Ejecución
# ---------------------------------------------------------------------------

def analyze_pdf(path: Path, backend: str):
    text = clean_text(open_pdfs(path))
    title = extract_company(text)
    sections = split_bonds_sections(text)
    conditions = extract_fields(text)

    print("=" * 70)
    print(path.name)
    if title:
        print(f"{title['company']} - Clase {title['bond']}")
    else:
        print("(no se pudo detectar empresa/clase)")
    print(f"Backend: {backend}")
    print("=" * 70)
    print()

    if backend == "local":
        print(build_cash_flow(conditions))
        return

    prompt = build_prompt(title, sections, conditions)
    try:
        print(ask_llm(prompt))
    except Exception as exc:
        print(f"[aviso] No se pudo usar el backend '{backend}': {exc}")
        print("[aviso] Se calcula el cash flow localmente como respaldo.")
        print()
        print(build_cash_flow(conditions))


def select_pdfs(pdfs):
    selection = PDF_INDEX.lower()
    if selection in {"all", "*", ""}:
        return pdfs

    if selection.isdigit():
        index = int(selection)
        if index >= len(pdfs):
            raise SystemExit(
                f"PDF_INDEX={index} no es válido; hay {len(pdfs)} PDFs: "
                + ", ".join(p.name for p in pdfs))
        return [pdfs[index]]

    path = Path(PDF_INDEX)
    if not path.is_absolute():
        path = directory / path
    if not path.exists():
        raise SystemExit(f"No existe el PDF: {path}")
    return [path]


def main():
    if not directory.is_dir():
        raise SystemExit(f"No existe la carpeta: {directory}")

    pdfs = get_pdfs(directory)
    if not pdfs:
        raise SystemExit(f"No hay archivos PDF en {directory}")

    for path in select_pdfs(pdfs):
        try:
            analyze_pdf(path, BACKEND)
        except Exception as exc:
            print(f"ERROR procesando {path.name}: {exc}")
        print()


if __name__ == "__main__":
    main()
