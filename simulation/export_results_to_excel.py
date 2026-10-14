import os
import re
import pandas as pd
import dispaset as ds

# ============================================================
# 1. CARTELLA DELLA SIMULAZIONE
# ============================================================

simulation_path = r'../Simulations/simulation_(2)_BE-NUC'

# ============================================================
# 2. LETTURA INPUT E RISULTATI
# ============================================================

print()
print("==============================================")
print("CARICAMENTO SIMULAZIONE UNLEASH (2)_BE-NUC")
print("==============================================")

inputs, results = ds.get_sim_results(path=simulation_path)

print("Input e risultati caricati correttamente.")
print("Numero di risultati:", len(results))

# ============================================================
# 3. FILE EXCEL DI OUTPUT
# ============================================================

output_file = os.path.join(
    simulation_path,
    'Risultati_completi_UNLEASH_BE_NUC.xlsx'
)

# ============================================================
# 4. FUNZIONI DI SUPPORTO
# ============================================================

used_sheet_names = set()


def clean_sheet_name(name, prefix=""):
    """
    Crea un nome di foglio Excel valido, massimo 31 caratteri
    e senza caratteri vietati.
    """

    name = prefix + str(name)

    # caratteri non ammessi in Excel
    name = re.sub(r'[\[\]\*\?/\\:]', '_', name)

    # massimo 31 caratteri
    name = name[:31]

    base = name
    counter = 1

    while name in used_sheet_names:
        suffix = f"_{counter}"
        name = base[:31 - len(suffix)] + suffix
        counter += 1

    used_sheet_names.add(name)

    return name


def prepare_dataframe(data):
    """
    Converte Series/DataFrame in un formato compatibile con Excel.
    """

    if isinstance(data, pd.Series):
        data = data.to_frame()

    if not isinstance(data, pd.DataFrame):
        return None

    df = data.copy()

    # Excel non supporta datetime con timezone
    if isinstance(df.index, pd.DatetimeIndex):
        if df.index.tz is not None:
            df.index = df.index.tz_localize(None)

    for col in df.columns:
        try:
            if pd.api.types.is_datetime64tz_dtype(df[col]):
                df[col] = df[col].dt.tz_localize(None)
        except Exception:
            pass

    return df


# ============================================================
# 5. CREAZIONE EXCEL
# ============================================================

index_rows = []

print()
print("Creazione file Excel...")
print()

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:

    # --------------------------------------------------------
    # FOGLIO SUMMARY
    # --------------------------------------------------------

    summary = pd.DataFrame({
        'Voce': [
            'Scenario',
            'Simulation directory',
            'File risultati GAMS',
            'Numero risultati Dispa-SET',
            'Descrizione'
        ],
        'Valore': [
            '(2)_BE-NUC',
            simulation_path,
            'Results.gdx',
            len(results),
            'UNLEASH Boundary Sector - Belgian system 2023'
        ]
    })

    summary.to_excel(
        writer,
        sheet_name='SUMMARY',
        index=False
    )

    used_sheet_names.add('SUMMARY')

    # --------------------------------------------------------
    # INPUT TABELLARI
    # --------------------------------------------------------

    print("------ ESPORTAZIONE INPUT ------")

    for key, value in inputs.items():

        df = prepare_dataframe(value)

        if df is None:
            continue

        sheet_name = clean_sheet_name(key, prefix='IN_')

        try:
            df.to_excel(
                writer,
                sheet_name=sheet_name
            )

            print(
                f"INPUT  {key}: "
                f"{df.shape[0]} righe x {df.shape[1]} colonne"
            )

            index_rows.append({
                'Tipo': 'INPUT',
                'Nome Dispa-SET': key,
                'Foglio Excel': sheet_name,
                'Righe': df.shape[0],
                'Colonne': df.shape[1]
            })

        except Exception as e:

            print(f"INPUT SALTATO {key}: {e}")

    # --------------------------------------------------------
    # RISULTATI
    # --------------------------------------------------------

    print()
    print("------ ESPORTAZIONE RISULTATI ------")

    for key, value in results.items():

        df = prepare_dataframe(value)

        if df is None:
            continue

        sheet_name = clean_sheet_name(key, prefix='OUT_')

        try:
            df.to_excel(
                writer,
                sheet_name=sheet_name
            )

            print(
                f"RESULT {key}: "
                f"{df.shape[0]} righe x {df.shape[1]} colonne"
            )

            index_rows.append({
                'Tipo': 'RISULTATO',
                'Nome Dispa-SET': key,
                'Foglio Excel': sheet_name,
                'Righe': df.shape[0],
                'Colonne': df.shape[1]
            })

        except Exception as e:

            print(f"RISULTATO SALTATO {key}: {e}")

    # --------------------------------------------------------
    # FOGLIO INDICE
    # --------------------------------------------------------

    index_df = pd.DataFrame(index_rows)

    index_df.to_excel(
        writer,
        sheet_name='INDEX',
        index=False
    )

# ============================================================
# 6. FINE
# ============================================================

print()
print("==============================================")
print("EXCEL CREATO CORRETTAMENTE")
print("==============================================")
print()
print(output_file)
print()
print("Il file contiene:")
print("- SUMMARY")
print("- INDEX")
print("- fogli IN_... per gli input tabellari")
print("- fogli OUT_... per i risultati")
print()