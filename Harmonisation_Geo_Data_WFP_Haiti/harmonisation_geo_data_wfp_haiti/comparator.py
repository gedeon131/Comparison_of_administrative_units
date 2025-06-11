from rapidfuzz import process, fuzz
import unicodedata
import re


def compare_exact(reference_df, other_df):
    """
    Exact match (case-sensitive).
    """
    ref_values = reference_df["Commune_ADM2"].unique()
    other_values = set(other_df["Commune_ADM2"].unique())

    matches = [(ref, ref) for ref in ref_values if ref in other_values]
    missing = [ref for ref in ref_values if ref not in other_values]
    extra = [val for val in other_values if val not in ref_values]

    return {
        "common": matches,
        "missing_in_other": missing,
        "extra_in_other": extra
    }


def compare_case_insensitive(reference_df, other_df):
    """
    Exact match ignoring case sensitivity.
    """
    ref_series = reference_df["Commune_ADM2"]
    other_series = other_df["Commune_ADM2"]

    ref_map = {val.lower(): val for val in ref_series}
    other_map = {val.lower(): val for val in other_series}

    ref_keys = set(ref_map.keys())
    other_keys = set(other_map.keys())

    matches = [(ref_map[k], other_map[k]) for k in ref_keys & other_keys]
    missing = [ref_map[k] for k in ref_keys - other_keys]
    extra = [other_map[k] for k in other_keys - ref_keys]

    return {
        "common": matches,
        "missing_in_other": missing,
        "extra_in_other": extra
    }


def compare_normalized(reference_df, other_df):
    """
    Exact match ignoring accents, punctuation, case and extra spaces.
    """
    def normalize_string(s):
        s = unicodedata.normalize("NFKD", s).encode("ASCII", "ignore").decode("utf-8")
        s = s.lower()
        s = re.sub(r"[-_,.;']", " ", s)
        s = re.sub(r"\s+", " ", s)
        return s.strip()

    ref_series = reference_df["Commune_ADM2"]
    other_series = other_df["Commune_ADM2"]

    ref_map = {normalize_string(val): val for val in ref_series}
    other_map = {normalize_string(val): val for val in other_series}

    ref_keys = set(ref_map.keys())
    other_keys = set(other_map.keys())

    matches = [(ref_map[k], other_map[k]) for k in ref_keys & other_keys]
    missing = [ref_map[k] for k in ref_keys - other_keys]
    extra = [other_map[k] for k in other_keys - ref_keys]

    return {
        "common": matches,
        "missing_in_other": missing,
        "extra_in_other": extra
    }


def compare_fuzzy(reference_df, other_df, threshold=90):
    """
    Fuzzy match using token_sort_ratio.
    Matches are returned when score >= threshold.
    """
    ref_values = reference_df["Commune_ADM2"].unique()
    other_values = other_df["Commune_ADM2"].unique()

    matches = []
    no_match = []
    matched_others = set()

    for ref in ref_values:
        match, score, _ = process.extractOne(ref, other_values, scorer=fuzz.token_sort_ratio)
        if score >= threshold:
            matches.append((ref, match, score))
            matched_others.add(match)
        else:
            no_match.append(ref)

    extra = [val for val in other_values if val not in matched_others]

    return {
        "common": matches,
        "missing_in_other": no_match,
        "extra_in_other": extra
    }
