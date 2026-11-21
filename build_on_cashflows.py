# -*- coding: utf-8 -*-
"""Genera los cash flows de las ONs de la planilla y los agrega a ON_cashflows.xlsx.

Reglas (consistentes con el libro existente):

* Capital de referencia USD 100 y convención 30/360.
* Interés = VR antes del pago * tasa nominal anual * días/360.
* Los datos de la planilla no traen amortización, así que se asume **bullet**
  (100 % del capital en la fecha de vencimiento).
* Se agregan únicamente los pagos con fecha >= hoy.
* Si el bono ya venció no se agrega (se informa).
* Si el ticker ya existe en el libro no se toca (idempotente).

Fechas: la planilla mezcla M/D/YYYY y D/M/YYYY. Acá quedan normalizadas a ISO
a mano; los casos dudosos están marcados con una nota.
"""

from __future__ import annotations

import argparse
import calendar
import shutil
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

HEADERS = [
    "Titulo", "Fecha de Pago", "Nro. Cupón", "VR",
    "Renta", "Amortización", "Renta + Amortización",
]
DATE_FORMAT = "DD/MM/YYYY"
CAPITAL = 100.0
TODAY = date.today()


@dataclass(frozen=True)
class Bond:
    pesos: str
    usd: str
    emisor: str
    vencimiento: date
    tasa: float
    frecuencia: int
    prox_pago: date
    nota: str = ""


def d(value: str) -> date:
    return datetime.strptime(value, "%Y-%m-%d").date()


def b(pesos, usd, emisor, venc, tasa, freq, prox, nota=""):
    return Bond(pesos, usd, emisor, d(venc), float(tasa), int(freq), d(prox), nota)


# ---------------------------------------------------------------------------
# Datos de la planilla (fechas normalizadas a ISO)
# ---------------------------------------------------------------------------
# Notas sobre fechas:
#   EAC3O / PECIO / OT41O / OT42O / VSCTD venían en D/M/YYYY y se convirtieron.
#   RUCAO: "3/30/2025" no alinea con el vencimiento 12/03/2028; se ancla al
#          vencimiento (12/03 y 12/09) y se marca.
#   OTS2O: el pago del 24/01/2025 no alinea con el vencimiento 27/10/2027
#          (24 vs 27); se conserva la serie del 24 y el último cupón se ajusta
#          al 27/10/2027.
#   STCEO: vencimiento "7/21/25"; queda fuera por estar vencido.
DATA = [
    b("AEC1O", "AEC1D", "AES",           "2025-07-14", "0.080", 2, "2025-01-14"),
    b("ARC1O", "ARC1D", "AES",           "2031-08-01", "0.085", 4, "2025-02-01", "ya existe ARC1D"),
    b("BYCHO", "BYCHD", "Galicia",       "2028-10-10", "0.083", 2, "2025-04-10"),
    b("CAC5O", "CAC5D", "Capex",         "2028-08-25", "0.093", 2, "2025-02-25"),
    b("CP17O", "CP17D", "Combustibles",  "2025-03-08", "0.095", 2, "2025-03-08"),
    b("CP32O", "CP32D", "Combustibles",  "2025-12-09", "0.060", 2, "2025-06-09"),
    b("CP34O", "CP34D", "Combustibles",  "2026-02-23", "0.070", 2, "2025-02-23"),
    b("CRCEO", "CRCED", "Celulosa",      "2025-06-04", "0.100", 4, "2025-03-04"),
    b("CRCJO", "CRCJD", "Celulosa",      "2028-05-16", "0.093", 4, "2025-02-16"),
    b("CRCLO", "CRCLD", "Celulosa",      "2026-02-08", "0.080", 4, "2025-02-08"),
    b("CS38O", "CS38D", "Cresud",        "2026-01-03", "0.080", 2, "2025-01-03"),
    b("DNC3O", "DNC3D", "Edenor",        "2026-11-22", "0.098", 2, "2025-05-22"),
    b("DNC5O", "DNC5D", "Edenor",        "2028-08-05", "0.095", 2, "2025-02-05"),
    b("DNC7O", "DNC7D", "Edenor",        "2029-10-24", "0.098", 2, "2025-04-24"),
    b("GNCXO", "GNCXD", "Geneia",        "2027-09-02", "0.088", 2, "2025-03-02"),
    b("IRCFO", "IRCFD", "IRSA",          "2028-06-22", "0.088", 2, "2024-12-22"),
    b("IRCGO", "IRCGD", "IRSA",          "2025-03-25", "0.080", 2, "2025-01-31"),
    b("IRCJO", "IRCJD", "IRSA",          "2027-02-28", "0.070", 2, "2025-02-28"),
    b("IRCOO", "IRCOD", "IRSA",          "2029-10-23", "0.073", 2, "2025-04-23"),
    b("LOC2O", "LOC2D", "Loma",          "2025-12-21", "0.065", 2, "2024-12-21"),
    b("LOC3O", "LOC3D", "Loma",          "2026-03-11", "0.075", 2, "2025-03-11"),
    b("MGC9O", "MGC9D", "Pampa",         "2026-12-08", "0.095", 2, "2025-06-08"),
    b("MGCHO", "MGCHD", "Pampa",         "2025-11-04", "0.050", 2, "2025-05-04"),
    b("MGCMO", "MGCMD", "Abanesi",       "2031-09-10", "0.080", 2, "2025-03-10"),
    b("MGCNO", "MGCND", "Abanesi",       "2028-10-04", "0.058", 2, "2025-04-04"),
    b("MR35O", "MR35D", "Abanesi",       "2025-02-08", "0.095", 2, "2025-02-08"),
    b("MRCAO", "MRCAD", "Abanesi",       "2027-12-01", "0.096", 2, "2025-06-01"),
    b("MRCLO", "MRCLD", "Abanesi",       "2025-01-27", "0.090", 2, "2025-01-27"),
    b("MTCGO", "MTCGD", "Abanesi",       "2026-06-30", "0.110", 4, "2024-12-30"),
    b("PN35O", "PN35D", "PAE",           "2029-09-29", "0.070", 2, "2025-03-29"),
    b("PN36O", "PN36D", "PAE",           "2031-11-10", "0.073", 2, "2025-05-10"),
    b("PNDCO", "PNDCD", "PAE",           "2027-04-30", "0.091", 2, "2025-04-30"),
    b("PNXCO", "PNXCD", "PAE",           "2027-04-30", "0.085", 2, "2025-04-30"),
    b("RCCJO", "RCCJD", "MSU",           "2027-10-09", "0.083", 2, "2025-04-09"),
    b("RUC7O", "RUC7D", "MSU",           "2025-01-12", "0.075", 2, "2025-01-12"),
    b("RUCAO", "RUCAD", "MSU",           "2028-03-12", "0.083", 2, "2025-03-30", "prox. pago no alinea; se ancla al vencimiento"),
    b("TLC1O", "TLC1D", "Telecom",       "2026-07-18", "0.080", 2, "2025-01-18"),
    b("TLC5O", "TLC5D", "Telecom",       "2025-08-06", "0.085", 2, "2025-02-06"),
    b("TLCMO", "TLCMD", "Telecom",       "2031-07-31", "0.095", 2, "2025-01-31", "ya existe TLCMD (con fechas 18/07)"),
    b("VSCFO", "VSCFD", "Vista",         "2025-11-10", "0.063", 2, "2025-05-10"),
    b("VSCPO", "VSCPD", "Vista",         "2029-05-03", "0.080", 2, "2025-05-03"),
    b("VSCRO", "VSCRD", "Vista",         "2031-10-10", "0.077", 2, "2025-04-10"),
    b("XMC1O", "XMC1D", "Exar",          "2027-11-11", "0.080", 2, "2025-05-11"),
    b("YCA6O", "YCA6P", "YPF",           "2025-07-28", "0.085", 2, "2025-01-28", "ticker USD termina en P en la planilla"),
    b("YMCHO", "YMCHD", "YPF",           "2026-02-12", "0.090", 4, "2025-02-12"),
    b("YMCIO", "YMCID", "YPF",           "2029-06-27", "0.090", 2, "2024-12-27"),
    b("YMCJO", "YMCJD", "YPF",           "2033-09-30", "0.070", 2, "2025-03-30"),
    b("YMCQO", "YMCQD", "YPF",           "2026-02-13", "0.050", 2, "2024-12-13"),
    b("OTS2O", "OTS2D", "Oil Tanking",   "2027-10-27", "0.070", 4, "2025-01-24", "24 vs 27; último cupón ajustado al vencimiento"),
    b("OTS3O", "OTS3D", "Oil Tanking",   "2028-11-01", "0.070", 4, "2025-08-01"),
    b("GYC4O", "GYC4D", "360 Energy",    "2027-10-30", "0.080", 4, "2025-04-30"),
    b("TLCOO", "TLCOD", "Telecom",       "2028-11-28", "0.070", 2, "2025-05-28"),
    b("SNSBO", "SNSBD", "San Miguel",    "2026-10-14", "0.090", 4, "2025-04-14"),
    b("STCEO", "STCED", "Grupo ST",      "2025-07-21", "0.070", 4, "2025-01-04", "prox. pago dudoso; vencido"),
    b("MR39O", "MR39D", "Albanesi",      "2031-11-01", "0.110", 2, "2025-05-01", "planilla: 'Semetral'"),
    b("RUCDD", "RUCDO", "MSU",           "2030-12-05", "0.098", 2, "2025-05-05", "planilla invierte sufijos"),
    b("CP36O", "CP36D", "Combustibles",  "2027-10-10", "0.065", 2, "2025-04-10", "planilla: 'Semetral'"),
    b("EAC3O", "EAC3D", "MSU",           "2028-12-20", "0.085", 2, "2025-06-20", "planilla: 20/12/2028 y 20/06/205"),
    b("OT41O", "OT41D", "Oil Tanking",   "2028-01-17", "0.070", 4, "2025-10-17", "fechas D/M"),
    b("OT42O", "OT42D", "Oil Tanking",   "2030-01-17", "0.080", 4, "2025-10-17", "fechas D/M"),
    b("PECIO", "PECID", "Aconcagua",     "2028-01-30", "0.090", 2, "2025-07-30", "fechas D/M"),
    b("VSCTD", "VSCTO", "Vista",         "2035-12-10", "0.076", 2, "2025-06-10", "fechas D/M; planilla invierte sufijos"),
]


def add_months(value: date, months: int) -> date:
    index = value.month - 1 + months
    year = value.year + index // 12
    month = index % 12 + 1
    return date(year, month, min(value.day, calendar.monthrange(year, month)[1]))


def days_30_360(start: date, end: date) -> int:
    return 360 * (end.year - start.year) + 30 * (end.month - start.month) \
        + (min(end.day, 30) - min(start.day, 30))


def build_schedule(bond: Bond):
    """Devuelve [(fecha, dias, cupon_n)] desde el primer cupón conocido."""
    step = 12 // bond.frecuencia
    dates: list[date] = []
    cur = bond.prox_pago
    while cur <= bond.vencimiento:
        dates.append(cur)
        cur = add_months(cur, step)

    if not dates:
        dates = [bond.vencimiento]
    elif dates[-1] != bond.vencimiento:
        gap = (bond.vencimiento - dates[-1]).days
        if gap <= 10:
            dates[-1] = bond.vencimiento
        else:
            # No alinea: reconstruir hacia atrás desde el vencimiento.
            dates = []
            cur = bond.vencimiento
            while cur > bond.prox_pago:
                dates.append(cur)
                cur = add_months(cur, -step)
            dates.sort()

    rows = []
    previous = add_months(dates[0], -step)
    for n, coupon in enumerate(dates, start=1):
        rows.append((coupon, days_30_360(previous, coupon), n))
        previous = coupon
    return rows


def build_cashflows(bond: Bond):
    """Filas (fecha, nro, vr, renta, amort, total) con flujos futuros."""
    out = []
    for coupon, days, n in build_schedule(bond):
        if coupon < TODAY:
            continue
        renta = round(CAPITAL * bond.tasa * days / 360.0, 4)
        amort = CAPITAL if coupon == bond.vencimiento else 0.0
        out.append((coupon, n, CAPITAL, renta, round(amort, 4), round(renta + amort, 4)))
    return out


def existing_symbols(ws) -> set[str]:
    symbols = set()
    for row in range(2, ws.max_row + 1):
        value = ws.cell(row=row, column=1).value
        if value:
            symbols.add(str(value).strip())
    return symbols


def ensure_headers(ws) -> None:
    for column, header in enumerate(HEADERS, start=1):
        cell = ws.cell(row=1, column=column)
        if cell.value is None:
            cell.value = header


def append_rows(ws, symbol: str, flows) -> int:
    start = ws.max_row + 1
    row = start
    for fecha, nro, vr, renta, amort, total in flows:
        values = [symbol, fecha, nro, vr, renta, amort, total]
        for column, value in enumerate(values, start=1):
            cell = ws.cell(row=row, column=column, value=value)
            if column == 2:
                cell.number_format = DATE_FORMAT
            elif column >= 4:
                cell.number_format = "0.0000"
        row += 1
    return row - start


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workbook", type=Path, default=Path(__file__).with_name("ON_cashflows.xlsx"))
    parser.add_argument("--dry-run", action="store_true", help="Mostrar sin escribir")
    parser.add_argument("--verbose", action="store_true", help="Imprimir cada schedule")
    args = parser.parse_args()

    wb = openpyxl.load_workbook(args.workbook)
    ws = wb["Sheet1"]

    if not args.dry_run:
        backup = args.workbook.with_name(
            f"{args.workbook.stem}_backup_{datetime.now():%Y%m%d_%H%M%S}{args.workbook.suffix}")
        shutil.copy2(args.workbook, backup)
        print(f"Backup: {backup.name}")

    ensure_headers(ws)
    current = existing_symbols(ws)

    added, skipped, duplicated = [], [], []
    seen_pairs = set()
    for bond in DATA:
        pair = (bond.pesos, bond.usd)
        if pair in seen_pairs:
            duplicated.append((bond.pesos, "duplicado en la planilla"))
            continue
        seen_pairs.add(pair)

        flows = build_cashflows(bond)
        if not flows:
            skipped.append((f"{bond.pesos}/{bond.usd}", "vencido o sin pagos futuros"))
            continue
        if bond.pesos in current or bond.usd in current:
            duplicated.append((f"{bond.pesos}/{bond.usd}", "ya existe en el libro"))
            continue

        if args.verbose:
            print(f"\n{bond.pesos}/{bond.usd} ({bond.emisor}) {bond.tasa:.1%} "
                  f"{bond.frecuencia} p/a - venc. {bond.vencimiento}")
            for fecha, _, _, renta, amort, total in flows:
                print(f"  {fecha:%d/%m/%Y}  renta {renta:>8.4f}  amort {amort:>7.1f}  total {total:>8.4f}")

        if not args.dry_run:
            for symbol in (bond.pesos, bond.usd):
                append_rows(ws, symbol, flows)
        added.append((f"{bond.pesos}/{bond.usd}", len(flows)))

    if not args.dry_run:
        wb.save(args.workbook)
        wb.close()

    print(f"\nTickers agregados: {len(added)} ({'dry-run' if args.dry_run else 'escrito'})")
    for symbol, count in added:
        print(f"  + {symbol:<16} {count:>3} pagos")
    if duplicated:
        print("\nNo se tocaron (ya existían o duplicados):")
        for symbol, reason in duplicated:
            print(f"  = {symbol:<16} {reason}")
    if skipped:
        print("\nOmitidos (vencidos / sin flujos futuros):")
        for symbol, reason in skipped:
            print(f"  - {symbol:<16} {reason}")

    flags = [(f"{x.pesos}/{x.usd}", x.nota) for x in DATA if x.nota]
    if flags:
        print("\nAdvertencias de datos:")
        for symbol, note in flags:
            print(f"  ! {symbol:<16} {note}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
