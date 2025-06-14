import pandas as pd
import os

# Build relative paths
base_path = "Harmonisation_Geo_Data_WFP_Haiti/data"
ocha_path = os.path.join(base_path, "Derniere_version_Officielle_Source_OCHA.xlsx")
less_path = os.path.join(base_path, "HTCO LESS destination locations.xlsx")
scope_path = os.path.join(base_path, "Administrative_Area_SCOPE.xlsx")
comet_path = os.path.join(base_path, "Administrative_Area_COMET.xlsx")

# Load raw data
df_ocha_raw = pd.read_excel(ocha_path, sheet_name="ADM3", header=0)
df_less_raw = pd.read_excel(less_path, sheet_name=0, header=0)
df_scope_raw = pd.read_excel(scope_path, sheet_name=0, header=0)
df_comet_raw = pd.read_excel(comet_path, sheet_name=0, header=1)

# Clean & rename columns
df_ocha = df_ocha_raw[["ADM1_EN", "ADM2_PCODE", "ADM2_EN", "ADM3_EN"]].rename(columns={
    "ADM1_EN": "Departement_ADM1",
    "ADM2_PCODE": "ADM2_PCODE",
    "ADM2_EN": "Commune_ADM2",
    "ADM3_EN": "Section_Communale_ADM3"
})

df_comet = df_comet_raw[["Breakdown 1", "Breakdown 2", "Point of Interest"]].rename(columns={
    "Breakdown 1": "Departement_ADM1",
    "Breakdown 2": "Commune_ADM2",
    "Point of Interest": "Section_Communale_ADM3"
})

df_scope = df_scope_raw[["Departement", "Commune", "Section Communale"]].rename(columns={
    "Departement": "Departement_ADM1",
    "Commune": "Commune_ADM2",
    "Section Communale": "Section_Communale_ADM3"
})

df_less = df_less_raw[["Loading Point description"]].rename(columns={
    "Loading Point description": "Commune_ADM2"
})

# Extract unique values from Commune_ADM2 for each source
df_ocha_ADM2 = pd.DataFrame(df_ocha["Commune_ADM2"].dropna().unique(), columns=["Commune_ADM2"])
df_comet_ADM2 = pd.DataFrame(df_comet["Commune_ADM2"].dropna().unique(), columns=["Commune_ADM2"])
df_scope_ADM2 = pd.DataFrame(df_scope["Commune_ADM2"].dropna().unique(), columns=["Commune_ADM2"])
df_less_ADM2 = pd.DataFrame(df_less["Commune_ADM2"].dropna().unique(), columns=["Commune_ADM2"])

# ADM2 + ADM3 combinations
df_ocha_complet = pd.DataFrame((df_ocha["Commune_ADM2"] + "_&_" + df_ocha["Section_Communale_ADM3"]).dropna().unique(), columns=["ADM2_&_ADM3"])
df_scope_complet = pd.DataFrame((df_scope["Commune_ADM2"] + "_&_" + df_scope["Section_Communale_ADM3"]).dropna().unique(), columns=["ADM2_&_ADM3"])

# Clean up the ",HT" suffix from each row in df_less_ADM2
df_less_ADM2_1 = df_less_ADM2.copy()
df_less_ADM2_1["Commune_ADM2"] = df_less_ADM2_1["Commune_ADM2"].str.replace(r",HT$", "", regex=True)

def export_to_excel(result, source_name, col_name="Commune_ADM2"):
    """
    Export result (common, missing, extra) into Excel file.
    Handles string safety and skips empty files.
    """
    output_dir = "output"
    os.makedirs(output_dir, exist_ok=True)

    if col_name == "ADM2_&_ADM3":
        filename = os.path.join(output_dir, f"Comparison_ADM2_&ADM3_OCHA_vs_{source_name}.xlsx")
    else:
        filename = os.path.join(output_dir, f"comparison_OCHA_vs_{source_name}.xlsx")

    def safe_dataframe(data, colname):
        return pd.DataFrame([str(x) for x in data], columns=[colname])

    common_data = result["common"]
    if common_data and isinstance(common_data[0], tuple):
        if len(common_data[0]) == 3:
            df_common = pd.DataFrame(common_data, columns=["OCHA", source_name, "Score"])
        else:
            df_common = pd.DataFrame(common_data, columns=["OCHA", source_name])
    else:
        df_common = safe_dataframe(common_data, col_name)

    df_missing = safe_dataframe(result["missing_in_other"], col_name)
    df_extra = safe_dataframe(result["extra_in_other"], col_name)

    if df_common.empty and df_missing.empty and df_extra.empty:
        print("⚠️ Skipping export: all result sets are empty.")
        return

    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        df_common.to_excel(writer, sheet_name="Common", index=False)
        df_missing.to_excel(writer, sheet_name="Missing", index=False)
        df_extra.to_excel(writer, sheet_name="Extra", index=False)

    print(f"Exported: {filename}")
