# -*- coding: utf-8 -*-
"""
Argentine corporate bonds / notes market data client and yield helpers.

Improved version of ``script finance.py``. Main differences:

* No global side effects (no ``os.chdir`` at import time, no work in module scope).
* Everything is organised as functions plus a ``main()`` entry point.
* Network layer has timeouts, retries and explicit error handling.
* Repeated "fetch JSON -> DataFrame" logic is factored into a single helper.
* IRR calculation is fixed: cash flows can be matched by date (XIRR) instead of
  assuming equally spaced coupons, and invalid inputs raise instead of returning NaN.
* Amortising bonds are handled correctly: the quoted price is scaled by the
  current residual nominal value (``VR``) instead of assuming 100.
* The payment frequency is inferred from the schedule instead of always
  assuming two coupons per year.
* Macaulay and modified duration, convexity and DV01 are computed with actual
  dates and the XIRR.
* Paths are handled with ``pathlib`` and can be supplied via CLI arguments.

Original author: emafl
"""

from __future__ import annotations

import argparse
import logging
from pathlib import Path
from typing import Iterable, Sequence

import numpy as np
import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

try:
    import numpy_financial as npf
except ModuleNotFoundError as exc:  # pragma: no cover - environment dependent
    raise SystemExit(
        "Missing dependency 'numpy-financial'. Install it with: pip install numpy-financial"
    ) from exc

logger = logging.getLogger(__name__)

# --------------------------------------------------------------------------- #
# Configuration
# --------------------------------------------------------------------------- #

DEFAULT_LIVE_URL = "https://data912.com/live/"
DEFAULT_HISTORICAL_URL = "https://data912.com/historical/"
DEFAULT_TIMEOUT = 10

# Instruments used by the original script.
TICKERS_USD = [
    "PN35D", "ARC1D", "MGCUO", "PLC7D", "PLC7O", "NPCDD", "MCC2D", "MCC3D",
    "SNSBD", "SNSBO", "GYC4D", "GN49D", "NPCDO", "MTC2D",
]
TICKERS_ARS = ["MGCUO", "PLC7O", "NPCDO", "MCC3O", "MTC2O"]
TICKERS_LONG = ["VSCXD", "PLC7D", "YM44D"]

# Column names expected in the cash-flow workbook.
CASHFLOW_COLUMN_MAP = {
    "Fecha de Pago": "fecha",
    "Renta + Amortización": "Pago",
    "Nro. Cupón": "N",
    "Titulo": "symbol",
    # Present in the workbook; VR is the outstanding nominal *before* each
    # payment, so it is 100 until the bond starts amortising.
    "VR": "vr",
}


# --------------------------------------------------------------------------- #
# Data client
# --------------------------------------------------------------------------- #


class ArgCorpClient:
    """Thin client for the data912 live/historical endpoints."""

    def __init__(
        self,
        url: str = DEFAULT_LIVE_URL,
        historical_url: str = DEFAULT_HISTORICAL_URL,
        timeout: float = DEFAULT_TIMEOUT,
        session: requests.Session | None = None,
    ) -> None:
        self.url = url.rstrip("/") + "/"
        self.historical_url = historical_url.rstrip("/") + "/"
        self.timeout = timeout
        self._session = session or self._build_session()

    @staticmethod
    def _build_session() -> requests.Session:
        session = requests.Session()
        retries = Retry(
            total=3,
            backoff_factor=0.5,
            status_forcelist=(429, 500, 502, 503, 504),
            allowed_methods=("GET",),
        )
        adapter = HTTPAdapter(max_retries=retries)
        session.mount("https://", adapter)
        session.mount("http://", adapter)
        session.headers.update({"User-Agent": "arg-corp-client/1.0"})
        return session

    def fetch_json(self, instrumento: str) -> object:
        """Fetch one instrument endpoint and return the decoded JSON payload."""
        instrumento = (instrumento or "").strip()
        if not instrumento:
            raise ValueError("instrumento must be a non-empty string")

        endpoint = self.url + instrumento
        logger.info("Fetching data from: %s", endpoint)

        try:
            response = self._session.get(endpoint, timeout=self.timeout)
        except requests.Timeout as exc:
            raise requests.Timeout(
                f"Request timed out after {self.timeout}s: {endpoint}"
            ) from exc
        except requests.ConnectionError as exc:
            raise requests.ConnectionError(f"Failed to connect to {endpoint}") from exc

        try:
            response.raise_for_status()
        except requests.HTTPError as exc:
            raise requests.HTTPError(
                f"Request failed with status {response.status_code}: {response.text[:500]}",
                response=response,
            ) from exc

        try:
            return response.json()
        except ValueError as exc:
            raise ValueError(f"Response from {endpoint} is not valid JSON") from exc

    def fetch_dataframe(self, instrumento: str) -> pd.DataFrame:
        """Fetch an endpoint and normalise the payload into a DataFrame."""
        payload = self.fetch_json(instrumento)
        if isinstance(payload, dict):
            payload = [payload]
        if not isinstance(payload, list):
            raise TypeError(
                f"Unexpected payload type for {instrumento!r}: {type(payload).__name__}"
            )
        df = pd.DataFrame(payload)
        logger.info("Loaded %d rows from %s", len(df), instrumento)
        return df


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #


def filter_by_symbol(df: pd.DataFrame, tickers: Iterable[str]) -> pd.DataFrame:
    """Return the rows of ``df`` whose ``symbol`` column is in ``tickers``."""
    if "symbol" not in df.columns:
        raise KeyError("DataFrame has no 'symbol' column")
    tickers = list(tickers)
    return df.loc[df["symbol"].isin(tickers)].copy()


# Backwards-compatible alias for the original helper name.
choose_tickets = filter_by_symbol


def load_cashflows(
    path: str | Path,
    sheet_name: str | int = 0,
    decimal: str = ",",
) -> pd.DataFrame:
    """Read and normalise the cash-flow workbook."""
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Cash-flow workbook not found: {path}")

    df = pd.read_excel(path, sheet_name=sheet_name, decimal=decimal)
    df = df.rename(columns=CASHFLOW_COLUMN_MAP)

    required = {"fecha", "Pago", "symbol"}
    missing = required - set(df.columns)
    if missing:
        raise KeyError(f"Missing expected columns in {path.name}: {sorted(missing)}")

    df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce", dayfirst=True)
    df["Pago"] = pd.to_numeric(df["Pago"], errors="coerce")
    if "vr" in df.columns:
        df["vr"] = pd.to_numeric(df["vr"], errors="coerce")

    invalid = df["fecha"].isna().sum() + df["Pago"].isna().sum()
    if invalid:
        logger.warning("Dropping %d row(s) with invalid date/payment values", invalid)
        df = df.dropna(subset=["fecha", "Pago"])

    return df.sort_values(["symbol", "fecha"]).reset_index(drop=True)


def current_vr(
    df: pd.DataFrame,
    fecha: object | None = None,
    default: float = 100.0,
) -> float:
    """Residual nominal value (``VR``) currently outstanding.

    The workbook stores in each row the nominal value *before* that payment.
    Therefore the VR of the next future payment is the residual still alive
    today. Bonds that have not started amortising return 100, and workbooks
    without a ``vr`` column fall back to ``default``.
    """
    if "vr" not in df.columns:
        return default

    data = df.dropna(subset=["fecha", "vr"]).sort_values("fecha")
    if data.empty:
        return default

    ref = pd.Timestamp(fecha) if fecha is not None else pd.Timestamp.today().normalize()
    future = data[data["fecha"] > ref]

    # If the bond already matured there is no residual left to buy.
    if future.empty:
        return 0.0

    vr = float(future["vr"].iloc[0])
    return vr if np.isfinite(vr) and vr > 0 else default


def infer_frequency(fechas: Iterable[object], default: int = 2) -> int:
    """Estimate coupons per year from the spacing of the payment dates."""
    dates = pd.to_datetime(pd.Series(list(fechas)), errors="coerce").dropna()
    dates = dates.sort_values().drop_duplicates()
    if len(dates) < 2:
        return default

    gaps = np.diff(dates.to_numpy(dtype="datetime64[D]").astype(np.int64))
    median_gap = float(np.median(gaps))
    if median_gap <= 0:
        return default

    candidates = np.array([1, 2, 3, 4, 6, 12], dtype=float)
    return int(candidates[np.argmin(np.abs(candidates - 365.0 / median_gap))])


# --------------------------------------------------------------------------- #
# Yield calculations
# --------------------------------------------------------------------------- #


def irr(
    cashflows: Sequence[float], frequency: int = 2
) -> tuple[float, float]:
    """Periodic IRR and its annualised equivalent.

    ``cashflows`` must start with the (negative) initial investment followed by
    the periodic payments.
    """
    if frequency <= 0:
        raise ValueError("frequency must be a positive integer")
    flows = np.asarray(cashflows, dtype=float)
    if flows.size < 2:
        raise ValueError("At least two cash flows are required")
    if not (np.any(flows > 0) and np.any(flows < 0)):
        raise ValueError("Cash flows must contain both positive and negative values")

    tir_periodo = float(npf.irr(flows))
    if not np.isfinite(tir_periodo):
        raise ValueError("IRR did not converge for the given cash flows")

    tir_anual = (1.0 + tir_periodo) ** frequency - 1.0
    return tir_periodo, tir_anual


def tir(
    df: pd.DataFrame,
    inversion_inicial: float = -100.0,
    frecuencia: int = 2,
    columna: str = "Pago",
) -> tuple[float, float]:
    """IRR from a DataFrame of periodic payments (original ``calcular_tir``)."""
    if columna not in df.columns:
        raise KeyError(f"DataFrame has no {columna!r} column")

    pagos = pd.to_numeric(df[columna], errors="coerce").dropna().to_numpy(dtype=float)
    if pagos.size == 0:
        raise ValueError("No valid payments to discount")

    return irr(np.concatenate(([float(inversion_inicial)], pagos)), frequency=frecuencia)


# ``calcular_tir`` kept as an alias so existing callers keep working.
calcular_tir = tir


def xirr(
    fechas: Iterable[object],
    flujos: Iterable[float],
    tol: float = 1e-10,
    max_iter: int = 200,
) -> float:
    """Money-weighted return using actual dates (365-day year).

    Unlike :func:`irr`, this does not assume equally spaced cash flows, which is
    the usual case for sovereign/corporate bond schedules with amortisation.
    """
    dates = pd.to_datetime(pd.Series(list(fechas)), errors="coerce")
    flows = np.asarray(list(flujos), dtype=float)

    if len(dates) != flows.size:
        raise ValueError("fechas and flujos must have the same length")
    if flows.size < 2:
        raise ValueError("At least two cash flows are required")
    if dates.isna().any():
        raise ValueError("fechas contains invalid dates")
    if not (np.any(flows > 0) and np.any(flows < 0)):
        raise ValueError("Cash flows must contain both positive and negative values")

    t = (dates - dates.iloc[0]).dt.days.to_numpy(dtype=float) / 365.0

    def npv(rate: float) -> float:
        return float(np.sum(flows / (1.0 + rate) ** t))

    # Find a sign change over a wide grid, then bisect.
    grid = np.concatenate(
        [
            np.linspace(-0.999, -0.01, 100),
            np.linspace(0.0, 1.0, 101),
            np.geomspace(1.0001, 1000.0, 100),
        ]
    )

    prev_rate, prev_value = float(grid[0]), npv(float(grid[0]))
    bracket: tuple[float, float, float] | None = None
    for rate in grid[1:]:
        rate = float(rate)
        value = npv(rate)
        if value == 0.0:
            return rate
        if prev_value * value < 0.0:
            bracket = (prev_rate, rate, prev_value)
            break
        prev_rate, prev_value = rate, value

    if bracket is None:
        raise ValueError("Could not bracket an XIRR solution")

    low, high, value_low = bracket
    for _ in range(max_iter):
        mid = 0.5 * (low + high)
        value_mid = npv(mid)
        if abs(value_mid) < tol or (high - low) < 1e-12:
            return mid
        if value_low * value_mid < 0.0:
            high = mid
        else:
            low, value_low = mid, value_mid

    return 0.5 * (low + high)


def _dated_flows(
    df: pd.DataFrame,
    inversion_inicial: float,
    columna_pago: str = "Pago",
    columna_fecha: str = "fecha",
    fecha_inversion: object | None = None,
) -> tuple[np.ndarray, np.ndarray]:
    """Build the dated cash-flow vector: purchase first, then each payment."""
    for column in (columna_pago, columna_fecha):
        if column not in df.columns:
            raise KeyError(f"DataFrame has no {column!r} column")

    data = df[[columna_fecha, columna_pago]].dropna().sort_values(columna_fecha)
    if data.empty:
        raise ValueError("No dated payments to discount")

    fechas = pd.to_datetime(data[columna_fecha]).to_numpy()
    flujos = data[columna_pago].to_numpy(dtype=float)

    compra = pd.Timestamp(fecha_inversion) if fecha_inversion is not None else pd.Timestamp.today().normalize()
    if compra > pd.Timestamp(fechas[0]):
        compra = pd.Timestamp(fechas[0])

    all_fechas = np.concatenate(([compra.to_datetime64()], fechas))
    all_flujos = np.concatenate(([float(inversion_inicial)], flujos))
    return all_fechas, all_flujos


def tir_por_fechas(
    df: pd.DataFrame,
    inversion_inicial: float = -100.0,
    columna_pago: str = "Pago",
    columna_fecha: str = "fecha",
    fecha_inversion: object | None = None,
) -> float:
    """XIRR-style annual return for a dated payment schedule.

    Parameters
    ----------
    inversion_inicial:
        Negative cash flow representing the purchase price.
    fecha_inversion:
        Date the investment is made. Defaults to today, clamped so it never
        falls after the first payment.
    """
    fechas, flujos = _dated_flows(
        df,
        inversion_inicial,
        columna_pago=columna_pago,
        columna_fecha=columna_fecha,
        fecha_inversion=fecha_inversion,
    )
    return xirr(fechas, flujos)


def duration(
    fechas: Iterable[object],
    flujos: Iterable[float],
    tasa: float,
) -> tuple[float, float]:
    """Macaulay and modified duration in years (actual/365).

    ``fechas`` must start with the valuation date and ``flujos`` with the
    (negative) purchase price. ``tasa`` is the effective annual yield, i.e. the
    XIRR, so the Macaulay duration is the PV-weighted average time and the
    modified duration is ``macaulay / (1 + tasa)``.
    """
    dates = pd.to_datetime(pd.Series(list(fechas)), errors="coerce")
    flows = np.asarray(list(flujos), dtype=float)

    if len(dates) != flows.size:
        raise ValueError("fechas and flujos must have the same length")
    if flows.size < 2:
        raise ValueError("At least two cash flows are required")
    if dates.isna().any():
        raise ValueError("fechas contains invalid dates")
    if not np.isfinite(tasa) or tasa <= -1.0:
        raise ValueError("tasa must be a finite effective annual rate greater than -100%")

    t = (dates - dates.iloc[0]).dt.days.to_numpy(dtype=float) / 365.0
    pv = flows / (1.0 + tasa) ** t

    # The price paid is the PV of the future payments. Using it (instead of the
    # net PV, which is ~0 at the IRR) avoids catastrophic cancellation.
    price = float(np.sum(pv[1:]))
    if price <= 0 or not np.isfinite(price):
        raise ValueError("Cannot compute duration: non-positive present value")

    macaulay = float(np.sum(t * pv) / price)
    modified = macaulay / (1.0 + tasa)
    return macaulay, modified


def convexity(
    fechas: Iterable[object],
    flujos: Iterable[float],
    tasa: float,
) -> float:
    """Convexity in years² using actual/365 dates and an effective annual yield.

    ``fechas``/``flujos`` follow the same convention as :func:`duration`, and
    ``tasa`` is the effective annual yield (XIRR). For a zero-coupon bond of
    maturity ``T`` this returns ``T * (T + 1) / (1 + tasa) ** 2``.
    """
    dates = pd.to_datetime(pd.Series(list(fechas)), errors="coerce")
    flows = np.asarray(list(flujos), dtype=float)

    if len(dates) != flows.size:
        raise ValueError("fechas and flujos must have the same length")
    if flows.size < 2:
        raise ValueError("At least two cash flows are required")
    if dates.isna().any():
        raise ValueError("fechas contains invalid dates")
    if not np.isfinite(tasa) or tasa <= -1.0:
        raise ValueError("tasa must be a finite effective annual rate greater than -100%")

    t = (dates - dates.iloc[0]).dt.days.to_numpy(dtype=float) / 365.0
    pv = flows / (1.0 + tasa) ** t

    price = float(np.sum(pv[1:]))
    if price <= 0 or not np.isfinite(price):
        raise ValueError("Cannot compute convexity: non-positive present value")

    return float(np.sum(t * (t + 1.0) * pv) / price / (1.0 + tasa) ** 2)


def duracion_por_fechas(
    df: pd.DataFrame,
    inversion_inicial: float,
    tasa: float,
    columna_pago: str = "Pago",
    columna_fecha: str = "fecha",
    fecha_inversion: object | None = None,
) -> tuple[float, float]:
    """Macaulay/modified duration for a dated payment schedule."""
    fechas, flujos = _dated_flows(
        df,
        inversion_inicial,
        columna_pago=columna_pago,
        columna_fecha=columna_fecha,
        fecha_inversion=fecha_inversion,
    )
    return duration(fechas, flujos, tasa)


# --------------------------------------------------------------------------- #
# Application flow
# --------------------------------------------------------------------------- #


def load_market(client: ArgCorpClient) -> dict[str, pd.DataFrame]:
    """Fetch every dataset the original script used."""
    names = ("arg_corp", "arg_bonds", "arg_options", "arg_notes", "arg_cedears")
    return {name: client.fetch_dataframe(name) for name in names}


def summarise(market: dict[str, pd.DataFrame]) -> None:
    corp = market["arg_corp"]
    corp_mep = (
        corp[corp["symbol"].astype(str).str.endswith("D")]
        .sort_values("v", ascending=False)
        .reset_index(drop=True)
    )

    bonds = market["arg_bonds"]
    bonds_usd = bonds[bonds["symbol"].astype(str).str.endswith("D")]

    logger.info("arg_corp rows:      %d (%d MEP)", len(corp), len(corp_mep))
    logger.info("arg_bonds USD rows: %d", len(bonds_usd))
    logger.info("arg_options rows:   %d", len(market["arg_options"]))
    logger.info("arg_notes rows:     %d", len(market["arg_notes"]))
    logger.info("arg_cedears rows:   %d", len(market["arg_cedears"]))

    logger.info("Checklist (USD):    %d instruments", len(filter_by_symbol(corp, TICKERS_USD)))
    logger.info("Checklist (ARS):    %d instruments", len(filter_by_symbol(corp, TICKERS_ARS)))
    logger.info("Long list:          %d instruments", len(filter_by_symbol(corp, TICKERS_LONG)))


COMPARISON_COLUMNS = [
    "symbol",
    "precio",
    "volumen",
    "vr",
    "inversion",
    "frecuencia",
    "pagos",
    "suma_pagos",
    "primera_fecha",
    "ultima_fecha",
    "tir_periodo_pct",
    "tir_anual_pct",
    "xirr_pct",
    "duracion_macaulay_anios",
    "duracion_modificada_anios",
    "convexidad_anios2",
    "dv01",
]


# Corporates live in ``arg_corp``, but public bonds (e.g. T15E7) only appear in
# ``arg_bonds``/``arg_notes``, so prices are collected from every endpoint.
PRICE_ENDPOINTS = ("arg_corp", "arg_bonds", "arg_notes")


def fetch_quotes(
    client: ArgCorpClient,
    endpoints: Sequence[str] = PRICE_ENDPOINTS,
) -> dict[str, dict[str, float]]:
    """Latest price (``c``) and traded volume (``v``) per symbol.

    Earlier endpoints take precedence, so a corporate ticker quoted in
    ``arg_corp`` is never overwritten by a same-named bond elsewhere.
    """
    quotes: dict[str, dict[str, float]] = {}
    for endpoint in endpoints:
        try:
            data = client.fetch_dataframe(endpoint)
        except requests.RequestException as exc:
            logger.warning("No se pudo leer %s: %s", endpoint, exc)
            continue
        if "symbol" not in data.columns or "c" not in data.columns:
            logger.warning("%s no tiene columnas 'symbol'/'c'", endpoint)
            continue

        has_volume = "v" in data.columns
        columns = ["symbol", "c"] + (["v"] if has_volume else [])
        for row in data[columns].itertuples(index=False, name=None):
            symbol, value = row[0], row[1]
            volume = row[2] if has_volume else float("nan")
            key = str(symbol)
            value = pd.to_numeric(value, errors="coerce")
            if key in quotes or pd.isna(value) or float(value) <= 0:
                continue
            volume = pd.to_numeric(volume, errors="coerce")
            quotes[key] = {
                "precio": float(value),
                "volumen": float(volume) if pd.notna(volume) else float("nan"),
            }
    return quotes


def fetch_prices(
    client: ArgCorpClient,
    endpoints: Sequence[str] = PRICE_ENDPOINTS,
) -> dict[str, float]:
    """Last traded price (``c``) per symbol, searched across the live endpoints."""
    return {symbol: quote["precio"] for symbol, quote in fetch_quotes(client, endpoints).items()}


def bond_metrics(
    series: pd.DataFrame,
    precio: float,
    vr: float | None = None,
    frecuencia: int | None = None,
    precio_original: bool = False,
    volumen: float | None = None,
) -> dict[str, object]:
    """Compute yield metrics for one bond from its full cash-flow schedule.

    ``precio_original=True`` treats the quoted price as per 100 of *original*
    nominal (the convention of the data912 feed when a bond has already
    amortised), instead of per 100 of residual nominal.
    """
    ref = pd.Timestamp.today().normalize()
    schedule = series[series["fecha"] > ref].sort_values("fecha")
    if schedule.empty:
        raise ValueError("sin flujos de caja futuros")

    residual = float(vr) if vr is not None else current_vr(series, fecha=ref)
    if residual <= 0:
        raise ValueError("sin valor residual")

    if precio_original:
        inversion = -abs(float(precio))
    else:
        inversion = -abs(float(precio)) * residual / 100.0
    freq = int(frecuencia) if frecuencia is not None else infer_frequency(schedule["fecha"])

    tir_periodo, tir_anual = tir(schedule, inversion_inicial=inversion, frecuencia=freq)
    try:
        xirr_frac = tir_por_fechas(schedule, inversion)
        xirr_anual = xirr_frac * 100.0
    except ValueError as exc:
        logger.warning("No se pudo calcular XIRR: %s", exc)
        xirr_frac = float("nan")
        xirr_anual = float("nan")

    # Risk metrics use actual dates and the effective annual yield (XIRR),
    # falling back to the annualised periodic IRR if XIRR could not be solved.
    tasa_efectiva = xirr_frac if np.isfinite(xirr_frac) else tir_anual
    try:
        fechas_flujos, flujos = _dated_flows(schedule, inversion)
        macaulay, modificada = duration(fechas_flujos, flujos, tasa_efectiva)
        convexidad = convexity(fechas_flujos, flujos, tasa_efectiva)
    except ValueError as exc:
        logger.warning("No se pudieron calcular duration/convexidad: %s", exc)
        macaulay = modificada = convexidad = float("nan")
    dv01 = modificada * abs(inversion) * 0.0001

    return {
        "symbol": str(series["symbol"].iloc[0]),
        "precio": round(abs(float(precio)), 2),
        "volumen": (
            int(round(float(volumen)))
            if volumen is not None and np.isfinite(float(volumen))
            else float("nan")
        ),
        "vr": round(residual, 2),
        "inversion": round(inversion, 2),
        "frecuencia": freq,
        "pagos": len(schedule),
        "suma_pagos": round(float(schedule["Pago"].sum()), 2),
        "primera_fecha": schedule["fecha"].iloc[0].date().isoformat(),
        "ultima_fecha": schedule["fecha"].iloc[-1].date().isoformat(),
        "tir_periodo_pct": round(tir_periodo * 100.0, 4),
        "tir_anual_pct": round(tir_anual * 100.0, 4),
        "xirr_pct": round(xirr_anual, 4),
        "duracion_macaulay_anios": round(macaulay, 4),
        "duracion_modificada_anios": round(modificada, 4),
        "convexidad_anios2": round(convexidad, 4),
        "dv01": round(dv01, 6),
    }


def run(
    cashflow_path: Path,
    symbol: str = "PLC7D",
    precio: float = -101.0,
    vr: float | None = None,
    frecuencia: int | None = None,
) -> dict[str, object]:
    """Fetch market data and compute the YTMs for one bond.

    ``precio`` is the quoted clean price per 100 of residual nominal value. It
    is scaled by the current ``VR`` so amortising bonds are not over-valued
    (this was the ARC1D bug: the price was applied to the whole original
    nominal even though the bond had already amortised ~17%).
    """
    client = ArgCorpClient()
    market = load_market(client)
    summarise(market)

    cashflows = load_cashflows(cashflow_path)
    series = cashflows[cashflows["symbol"] == symbol]
    if series.empty:
        logger.warning("No cash flows found for %s", symbol)
        return {}

    volumen = float("nan")
    corp = market["arg_corp"]
    if {"symbol", "v"} <= set(corp.columns):
        match = pd.to_numeric(
            corp.loc[corp["symbol"].astype(str) == symbol, "v"], errors="coerce"
        )
        if not match.empty and pd.notna(match.iloc[0]):
            volumen = float(match.iloc[0])

    try:
        result = bond_metrics(
            series, precio, vr=vr, frecuencia=frecuencia, volumen=volumen
        )
    except ValueError as exc:
        logger.warning("%s: %s", symbol, exc)
        return {}

    logger.info(
        "%s - precio %.2f%% x VR %.2f%% = inversion %.2f | frecuencia %d",
        symbol,
        result["precio"],
        result["vr"],
        result["inversion"],
        result["frecuencia"],
    )
    if np.isfinite(float(result["volumen"])):
        logger.info("%s - volumen: %d nominales", symbol, result["volumen"])
    logger.info("%s - TIR por período: %.4f%%", symbol, result["tir_periodo_pct"])
    logger.info("%s - TIR anual:       %.4f%%", symbol, result["tir_anual_pct"])
    if np.isfinite(result["xirr_pct"]):
        logger.info("%s - TIR por fechas:  %.4f%%", symbol, result["xirr_pct"])
    if np.isfinite(result["duracion_macaulay_anios"]):
        logger.info(
            "%s - Macaulay: %.4f anios | Modified: %.4f anios",
            symbol,
            result["duracion_macaulay_anios"],
            result["duracion_modificada_anios"],
        )
        logger.info(
            "%s - Convexidad: %.4f | DV01: %.6f por 100 nominal",
            symbol,
            result["convexidad_anios2"],
            result["dv01"],
        )

    return result


def compare_all(
    cashflow_path: Path,
    csv_path: Path | None = None,
    vr: float | None = None,
    frecuencia: int | None = None,
    fx: float | None = None,
    solo_usd: bool = False,
    precio_original: bool = False,
) -> pd.DataFrame:
    """Compute the yield table for every bond in the workbook.

    Prices come from the latest close across the live endpoints; bonds without
    a price are skipped with a warning. If ``csv_path`` is given the table is
    written there.

    Symbols whose ticker does not end in "D" trade in pesos while the cash-flow
    workbook is denominated in USD, so their price is divided by ``fx`` (pesos
    per dollar). Without ``fx`` those bonds are skipped instead of producing a
    nonsensical yield.
    """
    client = ArgCorpClient()
    quotes = fetch_quotes(client)
    cashflows = load_cashflows(cashflow_path)

    rows: list[dict[str, object]] = []
    skipped: list[tuple[str, str]] = []
    for symbol, series in cashflows.groupby("symbol", sort=True):
        if solo_usd and not str(symbol).upper().endswith("D"):
            continue
        quote = quotes.get(str(symbol))
        if quote is None:
            skipped.append((str(symbol), "sin precio en data912"))
            continue
        price = quote["precio"]
        if not str(symbol).upper().endswith("D") and price > 1000:
            # Corporates quoted in pesos show prices in the thousands (parity
            # with the USD series times the FX rate); USD clean prices stay
            # below ~200, so the magnitude tells the two apart.
            if not fx or fx <= 0:
                skipped.append(
                    (str(symbol), "precio en pesos; indicá --fx para convertirlo a USD")
                )
                continue
            price = price / fx
        try:
            rows.append(bond_metrics(
                series, price, vr=vr, frecuencia=frecuencia,
                precio_original=precio_original, volumen=quote["volumen"],
            ))
        except (ValueError, KeyError) as exc:
            skipped.append((str(symbol), str(exc)))

    for symbol, reason in skipped:
        logger.warning("Omitido %s: %s", symbol, reason)

    if not rows:
        raise ValueError("No se pudo calcular ningun bono")

    table = (
        pd.DataFrame(rows, columns=COMPARISON_COLUMNS)
        .sort_values("xirr_pct", ascending=False, na_position="last")
        .reset_index(drop=True)
    )
    print(table.to_string(index=False))

    if csv_path is not None:
        write_csv(table, csv_path)
    return table


def write_csv(table: pd.DataFrame, path: Path) -> Path:
    """Write the comparison table as UTF-8 CSV and return the resolved path."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    table.to_csv(path, index=False, encoding="utf-8-sig")
    return path.resolve()


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Fetch Argentine bond market data and compute bond yields."
    )
    parser.add_argument(
        "--cashflows",
        type=Path,
        default=Path.cwd() / "ON_cashflows.xlsx",
        help="Path to the ON cash-flow workbook (default: ./ON_cashflows.xlsx)",
    )
    parser.add_argument("--symbol", default="PLC7D", help="Bond symbol to analyse")
    parser.add_argument(
        "--precio",
        type=float,
        default=-101.0,
        help="Quoted price per 100 of residual nominal (negative or positive)",
    )
    parser.add_argument(
        "--vr",
        type=float,
        default=None,
        help="Residual nominal value per 100 (default: auto-detect from the workbook)",
    )
    parser.add_argument(
        "--frecuencia",
        type=int,
        default=None,
        choices=(1, 2, 3, 4, 6, 12),
        help="Coupons per year (default: infer from the payment schedule)",
    )
    parser.add_argument(
        "--todos",
        action="store_true",
        help="Calcular todos los bonos del Excel y mostrar la tabla comparativa",
    )
    parser.add_argument(
        "--solo-d",
        action="store_true",
        help="Calcular solo las series en dólares (tickers terminados en D)",
    )
    parser.add_argument(
        "--precio-original",
        action="store_true",
        help=(
            "Interpretar el precio como 'por 100 de nominal original' (data912 "
            "cuando el bono ya amortizó) en lugar de 'por 100 de nominal residual'"
        ),
    )
    parser.add_argument(
        "--fx",
        type=float,
        default=None,
        help=(
            "Pesos por dólar para convertir precios en ARS de tickers que no "
            "terminan en D (sin este dato se omiten)"
        ),
    )
    parser.add_argument(
        "--csv",
        type=Path,
        default=None,
        help="Guardar la tabla en CSV (con --todos, por defecto tir_bonos.csv)",
    )
    parser.add_argument(
        "--log-level",
        default="INFO",
        choices=("DEBUG", "INFO", "WARNING", "ERROR"),
        help="Logging verbosity",
    )
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    logging.basicConfig(
        level=getattr(logging, args.log_level),
        format="%(asctime)s %(levelname)-8s %(name)s: %(message)s",
    )
    try:
        if args.todos:
            csv_path = args.csv or Path.cwd() / "tir_bonos.csv"
            table = compare_all(
                args.cashflows,
                csv_path=csv_path,
                vr=args.vr,
                frecuencia=args.frecuencia,
                fx=args.fx,
                solo_usd=args.solo_d,
                precio_original=args.precio_original,
            )
            logger.info("Tabla guardada en %s (%d bonos)", csv_path.resolve(), len(table))
        else:
            result = run(
                args.cashflows,
                symbol=args.symbol,
                precio=args.precio,
                vr=args.vr,
                frecuencia=args.frecuencia,
            )
            if args.csv is not None and result:
                path = write_csv(pd.DataFrame([result], columns=COMPARISON_COLUMNS), args.csv)
                logger.info("Resultado guardado en %s", path)
    except (FileNotFoundError, KeyError, ValueError, requests.RequestException) as exc:
        logger.error("%s", exc)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
