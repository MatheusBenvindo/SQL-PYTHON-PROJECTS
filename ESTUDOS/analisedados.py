import pandas as pd
import numpy as np
from datetime import datetime, timedelta


# Função para calcular horas úteis
def calcular_horas_uteis(data_abertura, data_encerramento):
    # Verificar se há valores nulos
    if pd.isnull(data_abertura) or pd.isnull(data_encerramento):
        return timedelta(0)

    if data_abertura > data_encerramento:
        return timedelta(0)

    # Definições de horário comercial
    inicio_comercial = timedelta(hours=8, minutes=30)
    fim_comercial = timedelta(hours=17, minutes=30)
    horas_comerciais = fim_comercial - inicio_comercial

    # Ajustar horas do primeiro dia
    if data_abertura.time() < datetime.strptime("08:30:00", "%H:%M:%S").time():
        data_abertura = datetime.combine(
            data_abertura.date(), datetime.strptime("08:30:00", "%H:%M:%S").time()
        )
    elif data_abertura.time() > datetime.strptime("17:30:00", "%H:%M:%S").time():
        data_abertura += timedelta(days=1)
        data_abertura = datetime.combine(
            data_abertura.date(), datetime.strptime("08:30:00", "%H:%M:%S").time()
        )

    # Ajustar horas do último dia
    if data_encerramento.time() > datetime.strptime("17:30:00", "%H:%M:%S").time():
        data_encerramento = datetime.combine(
            data_encerramento.date(), datetime.strptime("17:30:00", "%H:%M:%S").time()
        )
    elif data_encerramento.time() < datetime.strptime("08:30:00", "%H:%M:%S").time():
        data_encerramento -= timedelta(days=1)
        data_encerramento = datetime.combine(
            data_encerramento.date(), datetime.strptime("17:30:00", "%H:%M:%S").time()
        )

    total_horas = timedelta(0)

    current_day = data_abertura
    while current_day.date() <= data_encerramento.date():
        if current_day.weekday() < 5:  # Dias úteis (segunda a sexta)
            if current_day.date() == data_abertura.date():
                total_horas += min(
                    horas_comerciais,
                    datetime.combine(current_day.date(), data_encerramento.time())
                    - current_day,
                )
            elif current_day.date() == data_encerramento.date():
                total_horas += min(
                    horas_comerciais,
                    data_encerramento
                    - datetime.combine(
                        current_day.date(),
                        datetime.strptime("08:30:00", "%H:%M:%S").time(),
                    ),
                )
            else:
                total_horas += horas_comerciais
        current_day += timedelta(days=1)

    return total_horas


# Caminho para o arquivo Excel
caminho_arquivo = r"C:\Users\matheus.ribeiro\OneDrive - Central das Cooperativas de Crédito e Economia do DF\Desktop\2024__.xlsm"

# Ler o arquivo Excel
df = pd.read_excel(caminho_arquivo, sheet_name="CONSOLIDADO")

# Calcular horas úteis para cada linha
df["Horas Úteis"] = df.apply(
    lambda row: calcular_horas_uteis(row["data_abertura"], row["data_encerramento"]),
    axis=1,
)

# Converter timedelta para o formato hh:mm:ss
df["Horas Úteis"] = df["Horas Úteis"].apply(lambda x: str(x))

# Salvar em um novo arquivo Excel
df.to_excel("Horas_Uteis_Calculadas.xlsx", index=False)
