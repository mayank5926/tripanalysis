import re
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font

BASE = Path(__file__).parent
INPUT_DIR = BASE / "Input Files"
OUTPUT = BASE / "All Trips.xlsx"
COLUMNS = ["SN", "Entry Time", "Exit Time", "Vehicle Name", "Duration (dd:hh:mm)"]
IGNORE = {"7_3_2026Wave ETP.xls"}  # hand-built cross-site sheet, not a raw export


def site_from_filename(name: str) -> str:
    stem = Path(name).stem
    stem = re.sub(r"^\d{1,2}_\d{1,2}_\d{4}", "", stem)  # leading date, e.g. 10_15_2026
    stem = re.sub(r"\s*\(\d+\)$", "", stem)  # trailing copy suffix, e.g. (1)
    return stem.strip()


def read_trips(path: Path) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, dtype=object)
    # Header row isn't always the first row (some exports have a "Site" line above it)
    header_idx = next(
        i for i, row in raw.iterrows() if {"SN", "Vehicle Name"} <= set(row.astype(str).str.strip())
    )
    df = raw.iloc[header_idx + 1 :].copy()
    df.columns = raw.iloc[header_idx].astype(str).str.strip()
    df = df[COLUMNS].dropna(how="all")
    df = df[df["SN"].notna() & (df["SN"].astype(str).str.strip() != "")]
    return df


def to_datetime(series: pd.Series) -> pd.Series:
    series = series.replace("", None)
    # Cells formatted as dates arrive as datetimes; unformatted ones as Excel serial numbers
    nums = pd.to_numeric(series.where(series.map(lambda v: isinstance(v, (int, float)))), errors="coerce")
    from_serial = pd.to_datetime(nums, unit="D", origin="1899-12-30")
    return pd.to_datetime(series, errors="coerce").fillna(from_serial).dt.round("s")


def main():
    frames = []
    for path in sorted(INPUT_DIR.glob("*.xls*")):
        if path.name in IGNORE:
            print(f"{path.name}: skipped")
            continue
        site = site_from_filename(path.name)
        df = read_trips(path)
        df["Site Name"] = site
        df["File Name"] = path.name
        frames.append(df)
        print(f"{path.name}: {len(df)} rows -> {site}")

    all_trips = pd.concat(frames, ignore_index=True)
    all_trips["SN"] = pd.to_numeric(all_trips["SN"]).astype(int)
    all_trips["Entry Time"] = to_datetime(all_trips["Entry Time"])
    all_trips["Exit Time"] = to_datetime(all_trips["Exit Time"])
    all_trips = all_trips.sort_values(["Vehicle Name", "Entry Time"], kind="stable", ignore_index=True)

    # Same trip can appear in several exports, so ignore File Name when comparing
    before = len(all_trips)
    dedup_key = all_trips.columns.drop("File Name")
    all_trips = all_trips.drop_duplicates(subset=dedup_key, ignore_index=True)
    print(f"Rows before dedup: {before}, after: {len(all_trips)}")

    # A trip = leaving Wave ETP (has exit time) followed by arriving at another site (has entry time)
    is_wave = all_trips["Site Name"] == "Wave ETP"
    trip_start = is_wave & all_trips["Exit Time"].notna()
    trip_end = ~is_wave & all_trips["Entry Time"].notna()
    pair = trip_start & trip_end.shift(-1, fill_value=False)
    all_trips["Trip"] = ""
    all_trips.loc[pair | pair.shift(1, fill_value=False), "Trip"] = "Trip Row"
    all_trips = all_trips[all_trips["Trip"] == "Trip Row"].reset_index(drop=True)

    # Summarise each trip on its Wave ETP row, using the next row as the destination
    nxt = all_trips.shift(-1)
    is_start = (
        (all_trips["Site Name"] == "Wave ETP")
        & all_trips["Exit Time"].notna()
        & (all_trips["Vehicle Name"] == nxt["Vehicle Name"])
    )
    # One row per trip
    start, dest = all_trips[is_start], nxt[is_start]
    trips = pd.DataFrame({
        "Trip Number": range(1, len(start) + 1),
        "Vehicle": start["Vehicle Name"],
        "Start Time": start["Exit Time"],
        "End Time": dest["Entry Time"],
        "Start Site": start["Site Name"],
        "End Site": dest["Site Name"],
    })
    # Store duration as a fraction of a day so Excel treats it as a real time value
    trips["Duration"] = (trips["End Time"] - trips["Start Time"]) / pd.Timedelta(days=1)

    # column: (width, number format, horizontal alignment)
    formats = {
        "A": (12, "0", "center"),                   # Trip Number
        "B": (13, "@", "left"),                     # Vehicle
        "C": (19, "dd-mmm-yyyy hh:mm:ss", "center"),  # Start Time
        "D": (19, "dd-mmm-yyyy hh:mm:ss", "center"),  # End Time
        "E": (12, "@", "left"),                     # Start Site
        "F": (19, "@", "left"),                     # End Site
        "G": (11, "[h]:mm:ss", "center"),           # Duration
    }
    with pd.ExcelWriter(OUTPUT, engine="openpyxl") as writer:
        trips.to_excel(writer, sheet_name="All Trips", index=False)
        ws = writer.sheets["All Trips"]
        for col, (width, number_format, align) in formats.items():
            ws.column_dimensions[col].width = width
            for cell in ws[col]:
                cell.alignment = Alignment(horizontal="center" if cell.row == 1 else align)
                if cell.row > 1:
                    cell.number_format = number_format
        for cell in ws[1]:
            cell.font = Font(bold=True)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    print(f"\nWrote {len(trips)} trips to {OUTPUT}")


if __name__ == "__main__":
    main()
