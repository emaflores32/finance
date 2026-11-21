# Contexto de la sesión — ONs argentinas (actualizado 2026-11-19)

Documento de traspaso: qué se hizo, qué se decidió y cómo retomar. Pensado para
levantar el trabajo en una sesión nueva sin releer todo el historial.

## 1. Entorno

- **Intérprete**: `C:\Users\emafl\anaconda3\python.exe` (3.13.9, con numpy, pandas,
  requests, numpy-financial, openpyxl, pdfplumber, matplotlib, scipy).
- El `python` del PATH es el stub de Microsoft Store y **no funciona**.
- Repo: `C:\Users\emafl\OneDrive\Desktop\repo`.
- PDFs de avisos: `C:\Users\emafl\OneDrive\Desktop\Personal\ONs`.

## 2. Archivos

| Archivo | Estado |
|---|---|
| `script_finance_improved.py` | Cliente data912 + TIR/duración/convexidad. Modificado (ver §3). |
| `pdf exercise.py` | Extractor de avisos de resultados en PDF. Modificado. |
| `build_on_cashflows.py` | **Nuevo**: genera y agrega cronogramas desde una planilla. Idempotente. |
| `curva_rendimientos.py` | **Nuevo**: curva TIR vs. plazo y vs. duration (Nelson-Siegel robusto). |
| `ON_cashflows.xlsx` | Libro maestro: **660 filas, 93 tickers**. Modificado. |
| `tir_bonos.csv` | Última corrida: **43 series D**, precios por 100 original. Regenerable. |
| `curva_rendimientos.png` / `.csv` | Figura y tabla de la curva. Regenerables. |
| `script finance.py` | Original, sin tocar; movido a `oldversion/` (tiene `os.chdir` y efectos globales). |

`build_on_cashflows.py` crea un backup con timestamp del libro antes de escribir
(no se versionan; los de esta sesión ya no están en el repo).

## 3. Cambios en `script_finance_improved.py`

- `fetch_prices()` recorre **arg_corp + arg_bonds + arg_notes** (antes solo
  arg_corp, por eso faltaba `T15E7`).
- `--fx` (pesos por dólar): convierte precios en ARS a USD para tickers sin
  sufijo "D" cuyo precio supere 1000 (heurística). Sin `--fx` se omiten.
- **`--solo-d`** (nuevo): calcula solo tickers terminados en `D` (series USD).
- **`--precio-original`** (nuevo): interpreta el precio como *por 100 de nominal
  original* en vez de *por 100 de residual*. Ver §7, es el criterio correcto para
  data912 cuando el bono ya amortizó.
- `compare_all()` acepta `fx=`, `solo_usd=`, `precio_original=`. `bond_metrics()`
  acepta `precio_original=`.

## 4. Cambios en `pdf exercise.py`

- `extract_fields()`: acepta numeración `1.` y `1)`, y descarta puntos seguidos de
  dígito para no confundir con montos (`70.000.000`).
- `clean_text()`: normaliza a **NFC** (acentos descompuestos) y limpia espacios.
- `find_field()` + `_accent_tolerant()`: toleran acentos y caracteres de reemplazo.
- `parse_fixed_rate()`: acepta "Tasa de Interés" y "Tasa Aplicable", con fallback
  al primer porcentaje del campo.
- Refactor: `compute_cash_flow()` devuelve filas; `build_cash_flow()` las formatea.

## 5. Estado del libro `ON_cashflows.xlsx`

- 1 hoja (`Sheet1`), columnas: Titulo, Fecha de Pago, Nro. Cupón, VR, Renta,
  Amortización, Renta + Amortización. Fechas como fecha real (DD/MM/YYYY);
  autofiltro `A1:G661`.
- **660 filas / 93 tickers**. Los 18 originales (ARC1D, IRCPD, MCC3D, MGCOD,
  NPCDO, PLC4D, PLC7D, PN41D, S16M6, T15E7, T30J6, TLCMD, TSC3D, TTCAD, VSCVD,
  VSCXD, YM34D, YM44O) quedaron **intactos**.
- **Carga masiva 19/11**: `build_on_cashflows.py` agregó **37 pares O/D (74
  tickers)** desde una planilla de 60 filas, solo con pagos futuros:
  BYCH, CAC5, CRCJ, DNC3, DNC5, DNC7, GNCX, IRCF, IRCJ, IRCO, MGC9, MGCM,
  MGCN, MRCA, PN35, PN36, PNDC, PNXC, RCCJ, RUCA, VSCP, VSCR, XMC1, YMCI,
  YMCJ, OTS2, OTS3, GYC4, TLCO, MR39, RUCD, CP36, EAC3, OT41, OT42, PECI, VSCT.
- Supuestos de la carga masiva: **bullet** (la planilla no traía amortización),
  capital USD 100, 30/360, `Nro. Cupón` desde el primer cupón conocido.
- **Duplicados no tocados**: ARC1D y TLCMD ya existían.
- **Vencidos omitidos** (23 pares, sin pagos futuros): AEC1, CP17, CP32, CP34,
  CRCE, CRCL, CS38, IRCG, LOC2, LOC3, MGCH, MR35, MRCL, MTCG, RUC7, TLC1, TLC5,
  VSCF, YCA6, YMCH, YMCQ, SNSB, STCE.
- **Flujos reales cargados** (reemplazan el bullet):
  - `IRCFO`/`IRCFD`: VR 47,5 → cupones 9-12; amortiza 17,5 (22/06/2027) y 30
    (22/06/2028). Renta 2,08 / 1,31.
  - `CAC5O`/`CAC5D`: VR 50 → cupones 7-10; amortiza 12,5 por semestre
    (25/02/2027 a 25/08/2028). Renta 2,33 / 1,72 / 1,17 / 0,58.
- `YM44D` (19/11): se agregó copiando el cronograma de `YM44O` (misma clase,
  17 pagos, 7,55 % semestral, bullet 18/06/2035). Antes faltaba y por eso no
  aparecía en el CSV filtrado a series D.
- **Ojo**: al guardar con openpyxl se perdieron los valores cacheados de 72
  fórmulas. El libro no tiene fórmulas, solo valores.

## 6. Última corrida (19/11/2026, series D, precios por 100 original)

```powershell
& C:\Users\emafl\anaconda3\python.exe script_finance_improved.py --todos --solo-d --precio-original
```

43 bonos. Top (XIRR): `PNDCD 3.443 %` y `CAC5D 65,8 %` **no usar** (amortización
previa no cargada), `YMCID 13,70 · DNC7D 12,84 · DNC3D 12,30 · EAC3D 11,57 ·
GYC4D 11,18 · RUCDD 9,37 · YM44D 8,66 · MCC3D 8,40 · DNC5D 8,40 · VSCXD 8,29`.

IRSA: `IRCOD 5,42 · IRCFD 3,96 · IRCJD 1,77`. CAC5D: 6,41. YM44D: 8,66 (precio
95,60). ARC1D: −1,85 (ver §7).

## 7. Decisiones y advertencias

- **Convención de precios (importante)**: data912 cotiza **por 100 de nominal
  original** cuando el bono ya amortizó. Evidencia: IRCFD precio 51,90 con VR
  47,5 y flujos reales → TIR 3,96 % (coherente con IRSA); escalando por VR daría
  ~47 %. Por eso se usa `--precio-original` (inversión = precio, sin escalar).
  El criterio viejo (precio × VR/100) queda solo como default histórico.
- **ARC1D**: con VR 82,64 y precio 104,70 da −1,85 % (120 % de paridad). El
  cronograma/VR del libro parece desactualizado (la amortización probablemente
  arranca un año después). **Revisar antes de usarlo.**
- **PNDCD/PNXCD**: PNDCD precio 21,46 con bullet 100 → absurdo (falta
  amortización previa). PNXCD precio 110 vs. pago final 104,25 → negativo;
  revisar VR/vencimiento.
- **DNC7D/YMCID**: precio ~94 y TIR 13 %; probable amortización parcial.
- **FX**: se usó CCL de dolarapi (≈ 1624,85). La TIR de series en pesos depende
  del FX; por eso el CSV ahora es solo D.
- **T15E7**: TIR 25 % es el anualizado de una letra a 58 días, no comparable.
- **Rotación NPCDO → MCC3D**: +1,22 pp con duration casi igual, pero MCC3 es ley
  local: es prima por jurisdicción, no arbitraje. No recomendada por 1,2 pp.

## 8. Curva de rendimientos y duration (19/11/2026)

`curva_rendimientos.py` lee `tir_bonos.csv`, agrega años al vencimiento y ajusta
**Nelson-Siegel** (robusto, `soft_l1`) en dos versiones: TIR vs. plazo y TIR vs.
duration modificada.

- Bonos usados: **37 de 43**. Excluidos por rango/plazo: PNDCD, PNXCD, DNC3D.
  Excluidos por datos dudosos (constante `SOSPECHOSOS`): YMCID, DNC7D, ARC1D.
  Se pueden incluir con `--incluir-dudosos`.
- Curva ajustada: 0,5 a 5,84 % · 1 a 5,96 % · 2 a 6,25 % · 3 a 6,57 % ·
  5 a 7,17 % · 7 a 7,66 % · 10 a 8,19 %. β0 = 9,81 %, τ = 2,23.
- Curva por duration: β0 = 12,6 %, τ = 3,31. Duración de 0,27 a 6,79 años.
- Dispersión alta (GYC4D/EAC3D ~11 % a ~1 año); la curva es indicativa.

## 9. Reproducir

```powershell
# Tabla completa de TIR (solo D, precio por 100 original)
& C:\Users\emafl\anaconda3\python.exe script_finance_improved.py --todos --solo-d --precio-original

# Un bono puntual
& C:\Users\emafl\anaconda3\python.exe script_finance_improved.py --symbol IRCFD --precio-original

# Curva de rendimientos y duration
& C:\Users\emafl\anaconda3\python.exe curva_rendimientos.py

# Regenerar la carga masiva (idempotente)
& C:\Users\emafl\anaconda3\python.exe build_on_cashflows.py --dry-run --verbose

# Procesar un aviso nuevo
$env:PDF_INDEX="YM44O.pdf"; & C:\Users\emafl\anaconda3\python.exe "pdf exercise.py"
```

## 10. Pendientes / próximos pasos

1. Cargar los flujos reales que faltan: `PNDCD`, `PNXCD`, `DNC7D`, `YMCID`
   (hoy distorsionan el ranking) y revisar `ARC1D`.
2. Incorporar la **ley aplicable** (local vs NY) de cada emisor: es el dato que
   más cambia el ranking de rotación y hoy no está.
3. Resolver `S16M6` y `T30J6` (sin precio en data912) o excluirlos del libro.
4. Evaluar si conviene reemplazar la heurística de `--fx` (precio > 1000) por un
   campo explícito de moneda por ticker.
5. Decidir si `--precio-original` pasa a ser el default (hoy es opt-in).
