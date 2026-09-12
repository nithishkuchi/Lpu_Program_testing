import csv

def load_programme_seed(path="lpu_monitor/config/programme_seed.csv") -> dict:
    seed = {}
    with open(path, encoding="utf-8") as f:
        for row in csv.DictReader(f):
            seed[row["OfficialCode"]] = row
    return seed