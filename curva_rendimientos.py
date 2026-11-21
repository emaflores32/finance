# -*- coding: utf-8 -*-
"""Curva de rendimientos y curva de duration a partir de tir_bonos.csv.

* Panel izquierdo: TIR (XIRR) vs. años al vencimiento + ajuste Nelson-Siegel.
* Panel derecho:  TIR (XIRR) vs. duration modificada + ajuste Nelson-Siegel.
* Guarda la figura y una tabla de la curva ajustada por plazos estándar.

Los bonos con datos no confiables (amortización previa no cargada, precios
fuera de rango o plazos de pocos días) se excluyen del ajuste y se informan.
"""

from __future__ import annotations

import argparse
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from scipy.optimize import least_squares

# Tickers apartados del ajuste: (motivo)
SOSPECHOSOS = {
    "PNDCD": "precio 21,5: amortización previa no cargada",
    "PNXCD": "precio 110 vs. pago final 104,25: revisar VR/vencimiento",
    "DNC3D": "vence en 3 días: TIR anualizada no comparable",
    "DNC7D": "precio 94,3: posible amortización parcial",
    "YMCID": "precio 94,4: posible amortización parcial",
    "ARC1D": "VR 82,64 vs. precio 104,7: cronograma a revisar",
}


def nelson_siegel(t: np.ndarray, b0: float, b1: float, b2: float, tau: float) -> np.ndarray:
    x = np.clip(t, 1e-6, None) / tau
    factor = (1.0 - np.exp(-x)) / x
    return b0 + b1 * factor + b2 * (factor - np.exp(-x))


def fit_ns(t: np.ndarray, y: np.ndarray) -> np.ndarray:
    start = np.array([np.nanmedian(y), -0.02, 0.0, 2.0])

    def residuals(params):
        return nelson_siegel(t, *params) - y

    result = least_squares(
        residuals,
        start,
        bounds=([0.0, -0.5, -0.5, 0.05], [0.60, 0.5, 0.5, 12.0]),
        loss="soft_l1",
        f_scale=0.005,
        max_nfev=20000,
    )
    return result.x


def cargar(path: Path, incluir_dudosos: bool) -> tuple[pd.DataFrame, pd.DataFrame]:
    df = pd.read_csv(path, encoding="utf-8-sig")
    hoy = pd.Timestamp.today().normalize()
    df["vencimiento"] = pd.to_datetime(df["ultima_fecha"])
    df["anios"] = (df["vencimiento"] - hoy).dt.days / 365.25

    validos = df[
        df["xirr_pct"].between(-5.0, 25.0)
        & (df["anios"] > 0.25)
        & (df["duracion_modificada_anios"] > 0.05)
        & df["xirr_pct"].notna()
    ].copy()

    motivo = validos["symbol"].map(SOSPECHOSOS)
    if incluir_dudosos:
        excluidos = pd.DataFrame(columns=["symbol", "motivo"])
        return validos.sort_values("anios"), excluidos
    excluidos = validos.loc[motivo.notna(), ["symbol"]].copy()
    excluidos["motivo"] = motivo[motivo.notna()]
    return validos.loc[motivo.isna()].sort_values("anios"), excluidos


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--csv", type=Path, default=Path("tir_bonos.csv"))
    parser.add_argument("--salida", type=Path, default=Path("curva_rendimientos.png"))
    parser.add_argument("--tabla", type=Path, default=Path("curva_rendimientos.csv"))
    parser.add_argument("--incluir-dudosos", action="store_true")
    args = parser.parse_args()

    datos, excluidos = cargar(args.csv, args.incluir_dudosos)
    if len(datos) < 5:
        raise SystemExit("Muy pocos bonos válidos para ajustar una curva")

    t = datos["anios"].to_numpy(dtype=float)
    d = datos["duracion_modificada_anios"].to_numpy(dtype=float)
    y = datos["xirr_pct"].to_numpy(dtype=float) / 100.0

    params_plazo = fit_ns(t, y)
    params_dur = fit_ns(d, y)

    plazos = np.array([0.5, 0.75, 1, 1.5, 2, 3, 4, 5, 7, 10])
    curva_plazos = nelson_siegel(plazos, *params_plazo) * 100.0

    tabla = pd.DataFrame({
        "plazo_anios": plazos,
        "tir_ajustada_pct": np.round(curva_plazos, 4),
    })
    tabla.to_csv(args.tabla, index=False, encoding="utf-8-sig")

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(15, 6.5))
    colores = plt.cm.viridis(np.linspace(0.05, 0.85, len(datos)))

    for ax, x, params, etiqueta in (
        (ax1, t, params_plazo, "Años al vencimiento"),
        (ax2, d, params_dur, "Duration modificada (años)"),
    ):
        ax.scatter(x, y * 100.0, c=colores, s=45, edgecolor="white", linewidth=0.6,
                   zorder=3, label="Bonos")
        for xi, yi, symbol in zip(x, y * 100.0, datos["symbol"]):
            ax.annotate(symbol, (xi, yi), xytext=(4, 3), textcoords="offset points",
                        fontsize=6.5, color="#333333")
        grid = np.linspace(max(0.1, x.min() * 0.8), x.max() * 1.05, 300)
        ax.plot(grid, nelson_siegel(grid, *params) * 100.0, color="#c0392b",
                linewidth=2.2, label="Nelson-Siegel", zorder=2)
        ax.set_xlabel(etiqueta)
        ax.set_ylabel("TIR (XIRR, %)")
        ax.set_title(f"TIR vs. {etiqueta.lower()}")
        ax.grid(alpha=0.25)
        ax.legend(loc="best", fontsize=9)
        ax.set_xlim(left=0)

    fig.suptitle(
        "Curva de rendimientos — ONs argentinas (series D, USD)"
        f"  ·  {len(datos)} bonos  ·  {pd.Timestamp.today():%d/%m/%Y}",
        fontsize=13, fontweight="bold")
    fig.tight_layout(rect=(0, 0, 1, 0.95))
    fig.savefig(args.salida, dpi=160, bbox_inches="tight")
    plt.close(fig)

    print(f"Bonos usados: {len(datos)}  |  excluidos: {len(excluidos)}")
    if not excluidos.empty:
        for row in excluidos.itertuples():
            print(f"  - {row.symbol:<8} {row.motivo}")
    print("\nNelson-Siegel (plazos):  beta0={:.4f}  beta1={:.4f}  beta2={:.4f}  tau={:.3f}"
          .format(*params_plazo))
    print("Nelson-Siegel (duration): beta0={:.4f}  beta1={:.4f}  beta2={:.4f}  tau={:.3f}"
          .format(*params_dur))
    print("\nCurva ajustada por plazo:")
    for plazo, tasa in zip(plazos, curva_plazos):
        print(f"  {plazo:>5.2f} a  {tasa:>6.2f} %")
    print(f"\nFigura: {args.salida.resolve()}")
    print(f"Tabla:  {args.tabla.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
