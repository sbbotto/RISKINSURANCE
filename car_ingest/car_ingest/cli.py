from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

import pyodbc

from .config import load_config
from .docx_parser import parse_docx
from .sql_server import build_connection_string, ensure_schema_objects, upsert_report



def _connect(cfg):
    return pyodbc.connect(build_connection_string(cfg), autocommit=False)



def ingest_one(path: Path) -> None:
    cfg = load_config()
    report = parse_docx(path)
    conn = _connect(cfg)
    try:
        ensure_schema_objects(conn, cfg)
        upsert_report(conn, cfg, report)
        print(f"Ingested: {path}")
    finally:
        conn.close()



def ingest_folder_once(folder: Path) -> int:
    cfg = load_config()
    count = 0
    conn = _connect(cfg)
    try:
        ensure_schema_objects(conn, cfg)
        for path in sorted(folder.glob("*.docx")):
            if path.name.startswith("~$"):
                continue
            report = parse_docx(path)
            upsert_report(conn, cfg, report)
            print(f"Ingested: {path.name}")
            count += 1
    finally:
        conn.close()
    return count



def watch_folder(folder: Path, interval_seconds: int) -> None:
    cfg = load_config()
    seen: dict[str, tuple[float, int]] = {}
    print(f"Watching {folder} every {interval_seconds}s")
    while True:
        conn = _connect(cfg)
        try:
            ensure_schema_objects(conn, cfg)
            for path in sorted(folder.glob("*.docx")):
                if path.name.startswith("~$"):
                    continue
                stat = path.stat()
                marker = (stat.st_mtime, stat.st_size)
                key = str(path.resolve())
                if seen.get(key) != marker:
                    report = parse_docx(path)
                    upsert_report(conn, cfg, report)
                    seen[key] = marker
                    print(f"Updated: {path.name}")
            conn.commit()
        except Exception as exc:
            conn.rollback()
            print(f"Error: {exc}", file=sys.stderr)
        finally:
            conn.close()
        time.sleep(interval_seconds)



def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Ingest Case Assessment Report DOCX files into SQL Server.")
    sub = parser.add_subparsers(dest="command", required=True)

    sub.add_parser("setup", help="Create SQL Server tables.")

    p_ingest = sub.add_parser("ingest", help="Ingest all DOCX files in the watch folder once.")
    p_ingest.add_argument("--folder", type=Path, default=None)

    p_one = sub.add_parser("ingest-one", help="Ingest a single DOCX file.")
    p_one.add_argument("path", type=Path)

    p_watch = sub.add_parser("watch", help="Continuously watch the folder for new or changed DOCX files.")
    p_watch.add_argument("--folder", type=Path, default=None)
    p_watch.add_argument("--interval", type=int, default=None)

    args = parser.parse_args(argv)
    cfg = load_config()

    if args.command == "setup":
        conn = _connect(cfg)
        try:
            ensure_schema_objects(conn, cfg)
            conn.commit()
            print("Database objects created.")
        finally:
            conn.close()
        return 0

    if args.command == "ingest":
        folder = args.folder or cfg.watch_folder
        ingest_folder_once(folder)
        return 0

    if args.command == "ingest-one":
        ingest_one(args.path)
        return 0

    if args.command == "watch":
        folder = args.folder or cfg.watch_folder
        interval = args.interval or cfg.poll_interval_seconds
        watch_folder(folder, interval)
        return 0

    return 1


if __name__ == "__main__":
    raise SystemExit(main())
