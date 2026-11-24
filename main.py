"""CLI entrypoint for the Rajasthan SEC Panchayat Samiti Sarpanch Scraper."""

from __future__ import annotations

import argparse
import json
import logging
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict

from excel_builder import build_all_outputs
from scraper import SecRajasthanScraper


def setup_logging(logs_dir: Path) -> logging.Logger:
    logs_dir.mkdir(parents=True, exist_ok=True)
    log_file = logs_dir / f"scraper_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

    logger = logging.getLogger("raj_sec_scraper")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S")

    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(fmt)
    logger.addHandler(sh)

    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    logger.info("Logging initialized -> %s", log_file)
    return logger


def load_config(config_path: Path) -> Dict[str, Any]:
    with config_path.open("r", encoding="utf-8") as f:
        return json.load(f)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Rajasthan State Election Commission Panchayat Samiti-wise Contesting Sarpanch Scraper"
    )
    parser.add_argument(
        "--config",
        default="config.json",
        help="Path to configuration JSON file (default: config.json)",
    )
    parser.add_argument(
        "--discover",
        action="store_true",
        help="Discover all Districts & Panchayat Samitis from SEC portal and initialize progress tracker.",
    )
    parser.add_argument(
        "--force-rediscover",
        action="store_true",
        help="Force re-discovery of Districts & Panchayat Samitis even if catalog.json exists.",
    )
    parser.add_argument(
        "--smoke-test",
        type=int,
        default=0,
        metavar="N",
        help="Run a smoke test scraping N pending Panchayat Samitis (e.g., --smoke-test 4) + export parity check.",
    )
    parser.add_argument(
        "--run",
        action="store_true",
        help="Run full resumable scraping for all remaining pending/failed Panchayat Samitis.",
    )
    parser.add_argument(
        "--max-items",
        type=int,
        default=None,
        help="Optional maximum number of Panchayat Samitis to process during --run.",
    )
    parser.add_argument(
        "--retry-empty",
        action="store_true",
        help="Also re-attempt Panchayat Samitis previously marked empty_on_portal.",
    )
    parser.add_argument(
        "--build-excel",
        action="store_true",
        help="Build combined Excel workbook (AMBER + all scraped PSs), error report, and summary.",
    )
    parser.add_argument(
        "--all",
        action="store_true",
        help="Run discovery (if needed), full resumable scrape, and build all final Excel & report deliverables.",
    )

    args = parser.parse_args()
    base_dir = Path(__file__).resolve().parent
    config = load_config(base_dir / args.config)
    logger = setup_logging(base_dir / config["paths"]["logs_dir"])

    if not any([args.discover, args.smoke_test > 0, args.run, args.build_excel, args.all]):
        parser.print_help()
        return 1

    scraper = SecRajasthanScraper(config=config, base_dir=base_dir, logger=logger)

    if args.discover:
        counts = scraper.discover_all(force_refresh=args.force_rediscover)
        logger.info("Discovery summary: %s", counts)

    if args.smoke_test > 0:
        counts = scraper.run_smoke_test(count=args.smoke_test)
        logger.info("Smoke test summary: %s", counts)
        if args.build_excel:
            build_all_outputs(config=config, base_dir=base_dir, logger=logger)
        return 0

    if args.run or args.all:
        counts = scraper.run_full(
            max_items=args.max_items,
            retry_failed=True,
            retry_empty=args.retry_empty,
            force_rediscover=args.force_rediscover,
        )
        logger.info("Scraping run summary: %s", counts)

    if args.build_excel or args.all:
        res = build_all_outputs(config=config, base_dir=base_dir, logger=logger)
        logger.info("Build summary: %s", json.dumps(res, indent=2))

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
