from lpu_monitor.config.fee_row_mapping_builder import build_fee_row_mapping, find_single_file

INCOMING_INTL_FEE = "lpu_monitor/incoming/international_fee_excel"

if __name__ == "__main__":
    proglist = find_single_file("lpu_monitor/incoming/programme_list")
    intl_fee = find_single_file(INCOMING_INTL_FEE)

    matched, status_counts, unmatched = build_fee_row_mapping(
        proglist_path=proglist,
        fee_path=intl_fee,
        out_csv="lpu_monitor/config/international_fee_row_mapping.csv",
    )

    print(f"Matched {matched}, unmatched: {len(unmatched)}")

    if unmatched:
        print("Needs review:", unmatched)