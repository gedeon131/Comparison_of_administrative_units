from harmonisation_geo_data_wfp_haiti import utils, comparator, config
from harmonisation_geo_data_wfp_haiti.utils import export_to_excel


def print_level_menu():
    print("\n📌 Select comparison level:")
    print("1. Commune only (ADM2)")
    print("2. Commune + Section Communale (ADM2 + ADM3)")
    print("0. Exit")


def print_method_menu():
    print("\n📌 Select comparison method:")
    print("1. Exact match (case-sensitive)")
    print("2. Case-insensitive match")
    print("3. Normalized match (ignore accents, spaces, case)")
    print("4. Fuzzy match (threshold =", config.FUZZY_MATCH_THRESHOLD, ")")
    print("0. Exit")


def print_result(source_name, result, fuzzy=False):
    print(f"\n🔍 OCHA vs {source_name.upper()}")
    if fuzzy:
        print("✅ Fuzzy matches       :", len(result['common']))
        print("🟥 No match found      :", len(result['missing_in_other']))
        print("🟦 Extra in", source_name.ljust(10), ":", len(result['extra_in_other']))
    else:
        print("✅ Common values       :", len(result['common']))
        print("🟥 Missing in", source_name.ljust(10), ":", len(result['missing_in_other']))
        print("🟦 Extra in", source_name.ljust(10), ":", len(result['extra_in_other']))


def run_comparisons(method, col_name, df_ref, sources):
    for name, df_other in sources:
        if method == 1:
            result = comparator.compare_exact(df_ref, df_other, col_name=col_name)
        elif method == 2:
            result = comparator.compare_case_insensitive(df_ref, df_other, col_name=col_name)
        elif method == 3:
            result = comparator.compare_normalized(df_ref, df_other, col_name=col_name)
        elif method == 4:
            result = comparator.compare_fuzzy(df_ref, df_other, threshold=config.FUZZY_MATCH_THRESHOLD, col_name=col_name)
        else:
            print("Invalid method.")
            return

        print_result(name, result, fuzzy=(method == 4))
        export_to_excel(result, source_name=name, col_name=col_name)


def main():
    while True:
        print_level_menu()
        try:
            level_choice = int(input("👉 Enter your comparison level: "))
        except ValueError:
            print("Please enter a valid number.")
            continue

        if level_choice == 0:
            print("Exiting...")
            break
        elif level_choice not in [1, 2]:
            print("Invalid option. Please choose 1 or 2.")
            continue

        print_method_menu()
        try:
            method = int(input("Enter comparison method: "))
        except ValueError:
            print("Please enter a valid number.")
            continue

        if method == 0:
            print("Exiting...")
            break
        elif method not in [1, 2, 3, 4]:
            print("Invalid method. Please choose 1 to 4.")
            continue

        if level_choice == 1:
            df_ref = utils.df_ocha_ADM2
            sources = [
                ("COMET", utils.df_comet_ADM2),
                ("SCOPE", utils.df_scope_ADM2),
                ("LESS", utils.df_less_ADM2_1),
            ]
            run_comparisons(method, "Commune_ADM2", df_ref, sources)

        elif level_choice == 2:
            df_ref = utils.df_ocha_complet
            sources = [
                ("SCOPE", utils.df_scope_complet),
            ]
            run_comparisons(method, "ADM2_&_ADM3", df_ref, sources)


if __name__ == "__main__":
    main()
