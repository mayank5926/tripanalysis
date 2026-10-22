"""Dump the Postgres tracker table of every vehicle in All Trips.xlsx.

Vehicle -> IMEI comes from Azure SQL (Vehicle.Name / Vehicle.DeviceImei); Postgres holds one
table per IMEI. Connection details are read from the Fleetelligence project's .env.

Each table is written as a plain SQL file (CREATE TABLE, COPY data, then indexes and
constraints), the same layout pg_dump produces. Restore with:
    psql "<connection>" -f "IMEI Dumps/<file>.sql"
"""

import os
import sys
from contextlib import closing
from pathlib import Path

import pandas as pd
import psycopg
import pymssql
from dotenv import load_dotenv
from psycopg import sql

BASE = Path(__file__).parent
TRIPS_FILE = BASE / "All Trips.xlsx"
OUT_DIR = BASE / "IMEI Dumps"

load_dotenv(Path(r"d:\code\Fleetelligence\.env"))
SCHEMA = os.environ.get("PGSCHEMA", "public")


def normalise(reg: str) -> str:
    return "".join(ch for ch in str(reg).upper() if ch.isalnum())


def vehicle_imeis(vehicles: list[str]) -> pd.DataFrame:
    """Active Azure SQL vehicles whose registration matches, ignoring spaces, dashes and case."""
    keys = [normalise(v) for v in vehicles]
    placeholders = ", ".join(["%s"] * len(keys))
    query = f"""
        SELECT v.Id AS vehicle_id, v.Name AS vehicle_db_name, v.DeviceImei AS imei, c.Name AS company,
               UPPER(REPLACE(REPLACE(REPLACE(v.Name, ' ', ''), '-', ''), '.', '')) AS reg_key
        FROM Vehicle v
        LEFT JOIN Company c ON c.Id = v.CompanyId
        WHERE v.IsDeleted = 0
          AND UPPER(REPLACE(REPLACE(REPLACE(v.Name, ' ', ''), '-', ''), '.', '')) IN ({placeholders})
    """
    conn = pymssql.connect(
        server=os.environ["AZURE_SQL_SERVER"],
        port=os.environ.get("AZURE_SQL_PORT", "1433"),
        database=os.environ["AZURE_SQL_DATABASE"],
        user=os.environ["AZURE_SQL_USER"],
        password=os.environ["AZURE_SQL_PASSWORD"],
        tds_version="7.4",
        login_timeout=30,
    )
    with closing(conn), closing(conn.cursor()) as cur:
        cur.execute(query, tuple(keys))
        found = pd.DataFrame(cur.fetchall(), columns=[d[0] for d in cur.description])
    trips = pd.DataFrame({"vehicle": vehicles, "reg_key": keys})
    return trips.merge(found, on="reg_key", how="left").drop(columns="reg_key")


def table_ddl(cur: psycopg.Cursor, table: str) -> tuple[str, list[str]]:
    """CREATE TABLE statement, plus the constraint and index statements to run after the data."""
    cur.execute(
        """
        SELECT a.attname, format_type(a.atttypid, a.atttypmod), a.attnotnull,
               pg_get_expr(d.adbin, d.adrelid)
        FROM pg_attribute a
        LEFT JOIN pg_attrdef d ON d.adrelid = a.attrelid AND d.adnum = a.attnum
        WHERE a.attrelid = %s::regclass AND a.attnum > 0 AND NOT a.attisdropped
        ORDER BY a.attnum
        """,
        (table,),
    )
    cols = []
    for name, col_type, not_null, default in cur.fetchall():
        col = f"    {sql.Identifier(name).as_string(cur)} {col_type}"
        if default is not None:
            col += f" DEFAULT {default}"
        if not_null:
            col += " NOT NULL"
        cols.append(col)
    create = f"CREATE TABLE {table} (\n" + ",\n".join(cols) + "\n);"

    cur.execute(
        "SELECT conname, pg_get_constraintdef(oid) FROM pg_constraint WHERE conrelid = %s::regclass",
        (table,),
    )
    after = [
        f"ALTER TABLE {table} ADD CONSTRAINT {sql.Identifier(name).as_string(cur)} {definition};"
        for name, definition in cur.fetchall()
    ]
    # Indexes not already created by a constraint
    cur.execute(
        """
        SELECT pg_get_indexdef(i.indexrelid) FROM pg_index i
        WHERE i.indrelid = %s::regclass
          AND NOT EXISTS (SELECT 1 FROM pg_constraint c WHERE c.conindid = i.indexrelid)
        """,
        (table,),
    )
    after += [f"{row[0]};" for row in cur.fetchall()]
    return create, after


def dump_table(conn: psycopg.Connection, imei: str, out_path: Path) -> int:
    table = sql.Identifier(SCHEMA, imei).as_string(conn)
    tmp_path = out_path.with_suffix(".sql.part")
    with conn.cursor() as cur, open(tmp_path, "wb") as f:
        create, after = table_ddl(cur, table)
        cur.execute(sql.SQL("SELECT count(*) FROM {}").format(sql.Identifier(SCHEMA, imei)))
        rows = cur.fetchone()[0]
        f.write(f"-- {table}: {rows} rows\n".encode())
        f.write(f"CREATE SCHEMA IF NOT EXISTS {sql.Identifier(SCHEMA).as_string(conn)};\n\n".encode())
        f.write(f"{create}\n\nCOPY {table} FROM stdin;\n".encode())
        copy_query = sql.SQL("COPY {} TO STDOUT").format(sql.Identifier(SCHEMA, imei))
        with cur.copy(copy_query) as copy:
            for chunk in copy:
                f.write(chunk)
        f.write(b"\\.\n\n")
        f.write("\n".join(after).encode() + b"\n")
    tmp_path.replace(out_path)
    return rows


def main() -> int:
    vehicles = sorted(pd.read_excel(TRIPS_FILE)["Vehicle"].dropna().unique())
    mapping = vehicle_imeis(vehicles)
    OUT_DIR.mkdir(exist_ok=True)
    mapping.to_csv(OUT_DIR / "vehicle_imei_map.csv", index=False)
    print(mapping.to_string(index=False))

    problems = mapping.groupby("vehicle")["imei"].nunique()
    for vehicle, n in problems.items():
        if n != 1:
            print(f"  SKIP    {vehicle}: {n} IMEIs found in Azure SQL")

    failed = 0
    with psycopg.connect(os.environ.get("DATABASE_URL", ""), autocommit=True) as conn:
        for vehicle, imei in mapping.dropna(subset=["imei"]).drop_duplicates("vehicle")[["vehicle", "imei"]].values:
            if problems[vehicle] != 1:
                continue
            out_path = OUT_DIR / f"{vehicle}_{imei}.sql"
            try:
                rows = dump_table(conn, str(imei), out_path)
                print(f"  ok      {vehicle} {imei}: {rows:,} rows -> {out_path.name}")
            except psycopg.Error as e:
                failed += 1
                print(f"  FAILED  {vehicle} {imei}: {str(e).strip().splitlines()[0]}")
    return 1 if failed or (problems != 1).any() else 0


if __name__ == "__main__":
    sys.exit(main())
