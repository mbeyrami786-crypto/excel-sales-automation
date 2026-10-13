"""Generate sample sales data into input/sales.xlsx."""

from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

from core import INPUT_FILE

PRODUCT_PRICES = {
    "Laptop": 800,
    "Mouse": 20,
    "Keyboard": 50,
    "Monitor": 250,
    "Headset": 100,
}
REGIONS = ["North", "South", "East", "West"]


def generate(rows: int = 100, seed: int | None = 42) -> pd.DataFrame:
    import random

    if seed is not None:
        random.seed(seed)

    data = []
    for _ in range(rows):
        product = random.choice(list(PRODUCT_PRICES))
        data.append(
            {
                "Product": product,
                "Region": random.choice(REGIONS),
                "Quantity": random.randint(1, 15),
                "Price": PRODUCT_PRICES[product],
            }
        )
    return pd.DataFrame(data)


def main() -> int:
    parser = argparse.ArgumentParser(description="Create sample sales data.")
    parser.add_argument("--rows", type=int, default=100, help="number of rows (default: 100)")
    parser.add_argument(
        "--output",
        type=Path,
        default=INPUT_FILE,
        help=f"destination .xlsx (default: {INPUT_FILE})",
    )
    args = parser.parse_args()

    df = generate(args.rows)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(args.output, index=False)
    print(f"{args.rows}-row sample data written to {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
