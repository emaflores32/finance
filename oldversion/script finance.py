# -*- coding: utf-8 -*-
"""
Created on Mon Apr  6 10:36:38 2026

@author: emafl
"""
import os
import requests
import pandas as pd
import logging
import numpy as np
import numpy_financial as npf


os.chdir(r'C:\Users\emafl\OneDrive\Desktop\PA')


logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class ArgCorpClient:
     
    DEFAULT_URL = "https://data912.com/live/"
    HISTORICAL_URL = "https://data912.com/historical/"
    
    def __init__(self, url= DEFAULT_URL, historical_url: str = HISTORICAL_URL ,timeout: int = 10):
        self.url = url.rstrip("/") + "/"
        self.historical_url = historical_url.rstrip("/") + "/"
        self.timeout = timeout
        self._session = requests.Session()
        

    def fetch_json(self, instrumento: str) -> dict:
        if not instrumento or not instrumento.strip():
            raise ValueError("instrumento must be a non-empty string")
        
        endpoint = self.url + instrumento.strip()
        logger.info(f"Fetching data from: {endpoint}")
        
        try:
            response = self._session.get(endpoint, timeout=self.timeout)
            response.raise_for_status()
            return response.json()
        
        except requests.HTTPError as e:
            raise requests.HTTPError(f"Request failed with status {response.status_code}: {response.text}") from e
        except requests.ConnectionError as e:
            raise requests.ConnectionError(f"Failed to connect to {endpoint}") from e
        except requests.Timeout as e:
            raise requests.Timeout(f"Request timed out after {self.timeout}s") from e


def choose_tickets(tickers: list[str], df: pd.DataFrame) -> pd.DataFrame:
    """Filter DataFrame by one or more ticker symbols."""
    return df[df['symbol'].isin(tickers)]
# Example usage
if __name__ == "__main__":
    client = ArgCorpClient()
    data = client.fetch_json("arg_corp")
    print(data)
    
data = pd.DataFrame(data)    
corp_mep = data[data['symbol'].str.endswith("D")]
corp_mep = corp_mep.sort_values('v', ascending = False).reset_index(drop = True)

data_bonds = client.fetch_json("arg_bonds")
data_bonds = pd.DataFrame(data_bonds)
data_bonds_usd = data_bonds[data_bonds['symbol'].str.endswith("D")]

options = client.fetch_json("arg_options")
options = pd.DataFrame(options)

#prueba calculo tir

checklist = choose_tickets(['PN35D','ARC1D','MGCUO','PLC7D','PLC7O','NPCDD','MCC2D','MCC3D','SNSBD','SNSBO','GYC4D','GN49D','NPCDO','MTC2D'],data)

checklist_pesos = choose_tickets(['MGCUO','PLC7O','NPCDO','MCC3O','MTC2O'],data)

letras = client.fetch_json("arg_notes")
letras = pd.DataFrame(letras)

cedears = client.fetch_json("arg_cedears")
cedears = pd.DataFrame(cedears)


print(len(letras))


cashflows = pd.read_excel('ON_cashflows.xlsx', decimal = ',' )
cashflows = cashflows.rename(columns = {'Fecha de Pago': 'fecha', 'Renta + Amortización':'Pago','Nro. Cupón':'N',
                                        'Titulo':'symbol'})
cashflows['fecha']= pd.to_datetime(cashflows['fecha'])



def calcular_tir(df, inversion_inicial=-103.55, frecuencia=4):
    """
    Calcula la TIR anualizada de un bono a partir de un DataFrame con la columna 'Pago'.
    
    Parámetros:
    df : pandas.DataFrame
        Debe contener la columna 'Pago' con los flujos de caja positivos.
    inversion_inicial : float
        Flujo inicial negativo (ejemplo: -100 si el precio de emisión es 100).
    frecuencia : int
        Número de períodos por año (ejemplo: 4 si los pagos son trimestrales).
    
    Retorna:
    tir_periodo : float
        TIR por período.
    tir_anual : float
        TIR anualizada.
    """
    # Construir lista de flujos
    flujos = [inversion_inicial] + df["Pago"].tolist()
    
    # Calcular TIR por período
    tir_periodo = npf.irr(flujos)
    
    # Anualizar según frecuencia
    tir_anual = (1 + tir_periodo)**frecuencia - 1
    
    return tir_periodo, tir_anual


long = choose_tickets(['VSCXD','PLC7D','YM44D'],data)

df = cashflows[cashflows['symbol'] == 'PLC7D']
df = df[df["fecha"] > pd.Timestamp.today()]


tir_periodo, tir_anual = calcular_tir(df, inversion_inicial=-101, frecuencia=2)
print("TIR por período:", tir_periodo)
print("TIR anual:", tir_anual)
    