from harmonisation_geo_data_wfp_haiti import utils, comparator, config
from harmonisation_geo_data_wfp_haiti.utils import export_to_excel


def print_menu():
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
        print("✅ Common communes     :", len(result['common']))
        print("🟥 Missing in", source_name.ljust(10), ":", len(result['missing_in_other']))
        print("🟦 Extra in", source_name.ljust(10), ":", len(result['extra_in_other']))


def run_pairwise_comparisons(method):
    # Load data
    df_ocha_ADM2 = utils.df_ocha_ADM2
    df_comet_ADM2 = utils.df_comet_ADM2
    df_scope_ADM2 = utils.df_scope_ADM2
    df_less_ADM2 = utils.df_less_ADM2_1  # already cleaned

    pairs = [
        ("COMET", df_comet_ADM2),
        ("SCOPE", df_scope_ADM2),
        ("LESS", df_less_ADM2),
    ]

    for name, other_df in pairs:
        if method == 1:
            result = comparator.compare_exact(df_ocha_ADM2, other_df)
        elif method == 2:
            result = comparator.compare_case_insensitive(df_ocha_ADM2, other_df)
        elif method == 3:
            result = comparator.compare_normalized(df_ocha_ADM2, other_df) 
        elif method == 4:
            result = comparator.compare_fuzzy(df_ocha_ADM2, other_df, threshold=config.FUZZY_MATCH_THRESHOLD)
        else:
            print("❌ Invalid method.")
            return

        print_result(name, result, fuzzy=(method == 4))
        export_to_excel(result, name)



def main():
    while True:
        print_menu()
        try:
            choice = int(input("👉 Enter your choice: "))
        except ValueError:
            print("⚠️ Please enter a valid number.")
            continue

        if choice == 0:
            print("👋 Exiting...")
            break
        elif choice in [1, 2, 3, 4]:
            run_pairwise_comparisons(choice)
        else:
            print("⚠️ Invalid option. Please choose between 0 and 4.")


if __name__ == "__main__":
    main()
