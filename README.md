# TripsAnalysis

Builds a Wave ETP trips report from geofence export files (Mathura Refinery, Meerut IOCL Depot, Wave ETP).

A trip is a Wave ETP visit with an exit time, followed by an entry at another site for the same vehicle.

## Usage

Put the `.xls` exports in `Input Files/`, then:

```
pip install -r requirements.txt
python trips_pipeline.py
```

Each step is saved to `Steps/` for checking:

| File | Contents |
|---|---|
| `01 Combined.xlsx` | All rows from all input files, with Site Name and File Name |
| `02 Deduplicated.xlsx` | Repeated events removed (SN and File Name ignored) |
| `03 Sorted.xlsx` | By vehicle, then entry time (exit time when there's no entry) |
| `04 Trip Rows Marked.xlsx` | Trip column plus debug columns explaining each match |
| `05 Trip Rows Only.xlsx` | Only the trip rows |
| `06 Trip Details.xlsx` | Trip Number, Vehicle, Start/End Time, Start/End Site, Duration |
| `07 Trips.xlsx` | One formatted row per trip |

## Tests

```
python -m pytest --cov=trips_pipeline
```

## Other scripts

- `combine_trips.py`: the earlier single-file version of the report (writes `All Trips.xlsx`).
- `dump_imei_tables.py`: dumps each trip vehicle's Postgres IMEI table as a restorable `.sql` file. Reads connection details from the Fleetelligence project's `.env`.
