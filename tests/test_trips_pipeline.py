import datetime as dt
from pathlib import Path

import openpyxl
import pandas as pd
import pytest

import trips_pipeline as p

T = pd.Timestamp


def row(vehicle, site, entry=None, exit_=None, sn=1, file="f.xls", duration="00:00:00"):
    return {
        "SN": sn,
        "Entry Time": T(entry) if entry else pd.NaT,
        "Exit Time": T(exit_) if exit_ else pd.NaT,
        "Vehicle Name": vehicle,
        "Duration (dd:hh:mm)": duration,
        "Site Name": site,
        "File Name": file,
    }


def frame(*rows):
    return pd.DataFrame(list(rows))


def write_export(path: Path, rows: list[list], site_line: str | None = None) -> Path:
    """Write an export-like sheet; optional 'Site' line above the header like the newer exports."""
    wb = openpyxl.Workbook()
    ws = wb.active
    if site_line:
        ws.append(["Site", site_line, None, None, None, None])
    ws.append(["SN", "Entry Time", "Exit Time", "Vehicle Name", "Duration (dd:hh:mm)", "Map"])
    for r in rows:
        ws.append(r)
    wb.save(path)
    return path


@pytest.fixture
def dirs(tmp_path, monkeypatch):
    inp, steps = tmp_path / "Input Files", tmp_path / "Steps"
    inp.mkdir()
    monkeypatch.setattr(p, "INPUT_DIR", inp)
    monkeypatch.setattr(p, "STEPS_DIR", steps)
    monkeypatch.setattr(p, "IGNORE", {"ignore me.xlsx"})
    return inp, steps


# ---------- site_from_filename ----------

@pytest.mark.parametrize(
    "name, site",
    [
        ("10_15_2026Mathura Refinery.xls", "Mathura Refinery"),
        ("10_15_2026Mathura Refinery (1).xls", "Mathura Refinery"),
        ("10_15_2026Mathura Refinery (12).xls", "Mathura Refinery"),
        ("7_3_2026Wave ETP.xls", "Wave ETP"),
        ("1_1_2026Meerut IOCL Depot (2).xlsx", "Meerut IOCL Depot"),
        ("Wave ETP.xls", "Wave ETP"),
        ("10_22_2026 Wave ETP .xls", "Wave ETP"),
    ],
)
def test_site_from_filename(name, site):
    assert p.site_from_filename(name) == site


def test_site_from_filename_keeps_numbers_inside_name():
    assert p.site_from_filename("10_15_2026Depot 2.xls") == "Depot 2"


# ---------- to_datetime ----------

def test_to_datetime_passes_datetimes_through():
    s = pd.Series([dt.datetime(2026, 9, 26, 5, 32, 6)], dtype=object)
    assert p.to_datetime(s).tolist() == [T("2026-09-26 05:32:06")]


def test_to_datetime_converts_excel_serials():
    s = pd.Series([46291.230625, 46291], dtype=object)
    assert p.to_datetime(s).tolist() == [T("2026-09-26 05:32:06"), T("2026-09-26")]


def test_to_datetime_handles_mixed_datetimes_and_serials():
    s = pd.Series([dt.datetime(2026, 9, 26, 5, 32, 6), 46291.230625], dtype=object)
    assert p.to_datetime(s).tolist() == [T("2026-09-26 05:32:06"), T("2026-09-26 05:32:06")]


def test_to_datetime_blanks_become_nat():
    s = pd.Series(["", None, float("nan")], dtype=object)
    assert p.to_datetime(s).isna().all()


def test_to_datetime_rounds_to_seconds():
    s = pd.Series([dt.datetime(2026, 9, 26, 5, 32, 6, 600000)], dtype=object)
    assert p.to_datetime(s).iloc[0] == T("2026-09-26 05:32:07")


# ---------- read_file ----------

def test_read_file_header_on_first_row(tmp_path):
    path = write_export(tmp_path / "a.xlsx", [[1, 46291.2, None, "GJ1", "00:00:00", None]])
    df = p.read_file(path)
    assert list(df.columns) == p.COLUMNS
    assert len(df) == 1


def test_read_file_header_below_site_line(tmp_path):
    path = write_export(
        tmp_path / "a.xlsx",
        [[1, 46291.2, None, "GJ1", "00:00:00", None], [2, 46291.3, 46291.4, "GJ1", "00:02:24", None]],
        site_line="Wave ETP",
    )
    df = p.read_file(path)
    assert df["SN"].tolist() == [1, 2]


def test_read_file_drops_blank_and_footer_rows(tmp_path):
    path = write_export(
        tmp_path / "a.xlsx",
        [[1, 46291.2, None, "GJ1", "00:00:00", None], [None] * 6, ["Generated on 15-Oct", None, None, None, None, None]],
    )
    assert p.read_file(path)["SN"].tolist() == [1]


def test_read_file_without_header_raises_clear_error(tmp_path):
    wb = openpyxl.Workbook()
    wb.active.append(["nothing", "useful"])
    path = tmp_path / "bad.xlsx"
    wb.save(path)
    with pytest.raises(ValueError, match="bad.xlsx"):
        p.read_file(path)


def test_read_file_missing_required_column_raises(tmp_path):
    wb = openpyxl.Workbook()
    wb.active.append(["SN", "Vehicle Name"])
    wb.active.append([1, "GJ1"])
    path = tmp_path / "short.xlsx"
    wb.save(path)
    with pytest.raises(KeyError):
        p.read_file(path)


# ---------- combine ----------

def test_combine_adds_site_and_file_and_skips_ignored(dirs):
    inp, _ = dirs
    write_export(inp / "10_15_2026Wave ETP (1).xlsx", [[1, 46291.2, 46291.5, " GJ1 ", "00:07:12", None]])
    write_export(inp / "ignore me.xlsx", [[1, 46291.2, None, "GJ9", "00:00:00", None]])
    (inp / "notes.txt").write_text("not an export")
    df = p.combine()
    assert len(df) == 1
    assert df.loc[0, "Site Name"] == "Wave ETP"
    assert df.loc[0, "File Name"] == "10_15_2026Wave ETP (1).xlsx"
    assert df.loc[0, "Vehicle Name"] == "GJ1"
    assert df.loc[0, "SN"] == 1
    assert df.loc[0, "Exit Time"] == T("2026-09-26 12:00:00")


def test_combine_empty_folder_raises(dirs):
    with pytest.raises(FileNotFoundError):
        p.combine()


# ---------- dedupe ----------

def test_dedupe_removes_same_event_from_different_exports():
    df = frame(
        row("GJ1", "Wave ETP", "2026-09-26 08:00", "2026-09-26 17:00", sn=1, file="10_15.xls"),
        row("GJ1", "Wave ETP", "2026-09-26 08:00", "2026-09-26 17:00", sn=5, file="10_22.xls"),
    )
    out = p.dedupe(df)
    assert len(out) == 1
    assert out.loc[0, "File Name"] == "10_15.xls"  # first occurrence kept


def test_dedupe_keeps_different_events():
    df = frame(
        row("GJ1", "Wave ETP", "2026-09-26 08:00", "2026-09-26 17:00"),
        row("GJ1", "Wave ETP", "2026-09-26 08:01", "2026-09-26 17:00"),
        row("GJ2", "Wave ETP", "2026-09-26 08:00", "2026-09-26 17:00"),
        row("GJ1", "Mathura Refinery", "2026-09-26 08:00", "2026-09-26 17:00"),
    )
    assert len(p.dedupe(df)) == 4


# ---------- sort_rows ----------

def test_sort_by_vehicle_then_entry():
    df = frame(
        row("GJ2", "Wave ETP", "2026-09-01"),
        row("GJ1", "Wave ETP", "2026-09-03"),
        row("GJ1", "Wave ETP", "2026-09-02"),
    )
    out = p.sort_rows(df)
    assert out["Vehicle Name"].tolist() == ["GJ1", "GJ1", "GJ2"]
    assert out["Entry Time"].tolist() == [T("2026-09-02"), T("2026-09-03"), T("2026-09-01")]


def test_sort_places_exit_only_rows_by_exit_time():
    df = frame(
        row("GJ1", "Wave ETP", "2026-09-05"),
        row("GJ1", "Meerut IOCL Depot", None, "2026-09-01"),
    )
    out = p.sort_rows(df)
    assert out["Site Name"].tolist() == ["Meerut IOCL Depot", "Wave ETP"]
    assert "_sort_time" not in out.columns


# ---------- mark_trips ----------

def reasons(df):
    return p.mark_trips(df)["Debug Reason"].tolist()


def test_mark_trips_basic_pair():
    df = frame(
        row("GJ1", "Wave ETP", "2026-09-26 08:00", "2026-09-26 17:00"),
        row("GJ1", "Mathura Refinery", "2026-09-27 07:00"),
        row("GJ1", "Mathura Refinery", "2026-09-27 07:05"),
    )
    out = p.mark_trips(df)
    assert out["Trip"].tolist() == ["Trip Row", "Trip Row", ""]
    assert out["Debug Reason"].tolist() == ["trip start", "trip end", ""]


def test_mark_trips_only_last_wave_visit_starts_trip():
    df = frame(
        row("GJ1", "Wave ETP", "2026-09-26 08:00", "2026-09-26 08:05"),
        row("GJ1", "Wave ETP", "2026-09-26 08:10", "2026-09-26 17:00"),
        row("GJ1", "Meerut IOCL Depot", "2026-09-26 23:00", "2026-09-27 05:00"),
    )
    out = p.mark_trips(df)
    assert out["Trip"].tolist() == ["", "Trip Row", "Trip Row"]
    assert out.loc[0, "Debug Reason"] == "next row is also Wave ETP"


def test_mark_trips_wave_without_exit_is_not_a_trip():
    df = frame(row("GJ1", "Wave ETP", "2026-09-26 08:00"), row("GJ1", "Mathura Refinery", "2026-09-27"))
    assert reasons(df)[0] == "Wave ETP row has no exit time"
    assert p.mark_trips(df)["Trip"].tolist() == ["", ""]


def test_mark_trips_next_row_without_entry_is_not_a_trip():
    df = frame(
        row("GJ1", "Wave ETP", "2026-09-26 08:00", "2026-09-26 17:00"),
        row("GJ1", "Mathura Refinery", None, "2026-09-27"),
    )
    assert reasons(df)[0] == "next row has no entry time"


def test_mark_trips_last_row_in_file():
    df = frame(row("GJ1", "Wave ETP", "2026-09-26 08:00", "2026-09-26 17:00"))
    assert reasons(df) == ["last row in file"]


def test_mark_trips_last_row_for_vehicle_before_another_wave_row():
    df = frame(
        row("GJ1", "Wave ETP", "2026-09-26 08:00", "2026-09-26 17:00"),
        row("GJ2", "Wave ETP", "2026-09-20 08:00", "2026-09-20 17:00"),
    )
    assert reasons(df)[0] == "last row for this vehicle"


def test_mark_trips_flags_pair_across_vehicles():
    df = frame(
        row("GJ1", "Wave ETP", "2026-09-26 08:00", "2026-09-26 17:00"),
        row("GJ2", "Mathura Refinery", "2026-09-01 07:00"),
    )
    out = p.mark_trips(df)
    assert out["Trip"].tolist() == ["Trip Row", "Trip Row"]
    assert out.loc[0, "Debug Reason"] == "trip start - but next row is a DIFFERENT vehicle"
    assert out.loc[0, "Same Vehicle As Next"] == False  # noqa: E712


def test_mark_trips_non_wave_rows_never_start_trips():
    df = frame(
        row("GJ1", "Mathura Refinery", "2026-09-26 08:00", "2026-09-26 17:00"),
        row("GJ1", "Meerut IOCL Depot", "2026-09-27 07:00"),
    )
    assert p.mark_trips(df)["Trip"].tolist() == ["", ""]


def test_mark_trips_does_not_modify_input():
    df = frame(row("GJ1", "Wave ETP", "2026-09-26 08:00", "2026-09-26 17:00"))
    p.mark_trips(df)
    assert "Trip" not in df.columns


# ---------- trip_details ----------

def test_trip_details_fills_start_row_and_numbers_both_rows():
    df = frame(
        row("GJ1", "Wave ETP", "2026-09-26 08:00", "2026-09-26 17:30"),
        row("GJ1", "Mathura Refinery", "2026-09-27 07:45"),
        row("GJ2", "Wave ETP", "2026-09-28 08:00", "2026-09-28 19:00"),
        row("GJ2", "Meerut IOCL Depot", "2026-09-29 00:30", "2026-10-01 16:00"),
    )
    out = p.trip_details(df)
    assert out["Trip Number"].tolist() == [1, 1, 2, 2]
    start = out.loc[0]
    assert start["Vehicle"] == "GJ1"
    assert start["Start Time"] == T("2026-09-26 17:30")
    assert start["End Time"] == T("2026-09-27 07:45")
    assert start["Start Site"] == "Wave ETP"
    assert start["End Site"] == "Mathura Refinery"
    assert start["Duration"] == pd.Timedelta(hours=14, minutes=15)
    assert pd.isna(out.loc[1, "Vehicle"])
    assert out.loc[3, "End Site"] is pd.NA or pd.isna(out.loc[3, "End Site"])


def test_trip_details_ignores_pair_across_vehicles():
    df = frame(
        row("GJ1", "Wave ETP", "2026-09-26 08:00", "2026-09-26 17:00"),
        row("GJ2", "Mathura Refinery", "2026-09-27 07:00"),
    )
    out = p.trip_details(df)
    assert out["Vehicle"].isna().all()
    assert out["Trip Number"].isna().all()


# ---------- save ----------

def test_save_writes_formatted_file(dirs):
    _, steps = dirs
    steps.mkdir()
    df = pd.DataFrame({"Trip Number": [1], "Duration": [pd.Timedelta(hours=5, minutes=39)]})
    p.save(df, "x.xlsx", {"Duration": "[h]:mm:ss"})
    ws = openpyxl.load_workbook(steps / "x.xlsx").active
    assert ws["A1"].value == "Trip Number"
    assert ws["A1"].font.bold
    assert ws["B2"].number_format == "[h]:mm:ss"
    assert ws.freeze_panes == "A2"
    assert ws.auto_filter.ref == "A1:B2"


# ---------- main (end to end) ----------

def test_main_end_to_end(dirs):
    inp, steps = dirs
    # 15 Oct export
    write_export(
        inp / "10_15_2026Wave ETP.xlsx",
        [
            [1, 46291.0, 46291.2, "GJ1", "00:04:48", None],    # 26-Sep 00:00 - 04:48
            [2, 46291.3, 46291.7, "GJ1", "00:09:36", None],    # 26-Sep 07:12 - 16:48  -> trip start
            [3, 46293.3, 46293.8, "UP1", "00:12:00", None],    # 28-Sep 07:12 - 19:12  -> trip start
        ],
    )
    write_export(
        inp / "10_15_2026Mathura Refinery.xlsx",
        [[1, 46292.3, None, "GJ1", "00:00:00", None], [2, 46292.31, 46292.4, "GJ1", "00:00:14", None]],
        site_line="Mathura Refinery",
    )
    # 22 Oct export repeats the Wave rows with different SNs, and adds UP1's destination
    write_export(
        inp / "10_22_2026Wave ETP.xlsx",
        [[7, 46291.3, 46291.7, "GJ1", "00:09:36", None], [8, 46293.3, 46293.8, "UP1", "00:12:00", None]],
    )
    write_export(inp / "10_22_2026Meerut IOCL Depot (1).xlsx", [[1, 46294.0, 46296.5, "UP1", "02:12:00", None]])
    write_export(inp / "ignore me.xlsx", [[1, 46291.0, 46291.9, "ZZ9", "00:21:36", None]])

    p.main()

    names = sorted(f.name for f in steps.iterdir())
    assert names == [
        "01 Combined.xlsx",
        "02 Deduplicated.xlsx",
        "03 Sorted.xlsx",
        "04 Trip Rows Marked.xlsx",
        "05 Trip Rows Only.xlsx",
        "06 Trip Details.xlsx",
        "07 Trips.xlsx",
    ]
    assert len(pd.read_excel(steps / "01 Combined.xlsx")) == 8
    assert len(pd.read_excel(steps / "02 Deduplicated.xlsx")) == 6
    marked = pd.read_excel(steps / "04 Trip Rows Marked.xlsx")
    assert "Debug Reason" in marked.columns
    assert "Debug Reason" not in pd.read_excel(steps / "05 Trip Rows Only.xlsx").columns

    trips = pd.read_excel(steps / "07 Trips.xlsx")
    assert list(trips.columns) == ["Trip Number", "Vehicle", "Start Time", "End Time", "Start Site", "End Site", "Duration"]
    assert trips["Trip Number"].tolist() == [1, 2]
    assert trips["Vehicle"].tolist() == ["GJ1", "UP1"]
    assert trips["End Site"].tolist() == ["Mathura Refinery", "Meerut IOCL Depot"]
    assert trips.loc[0, "Start Time"] == T("2026-09-26 16:48")
    assert trips.loc[0, "End Time"] == T("2026-09-27 07:12")
