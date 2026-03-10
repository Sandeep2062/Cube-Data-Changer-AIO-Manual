"""
Cube Data Generator Module
Generates synthetic concrete and mortar cube test data (weights + strengths)
in-memory, ready for direct processing without intermediate Excel files.

Based on: https://github.com/Sandeep2062/Cube-Data-Generator
"""

import numpy as np


# ── Range definitions ───────────────────────────────────────────────────────

CONCRETE_GRADES = ["M10", "M15", "M20", "M25", "M30", "M35", "M40", "M45"]

MORTAR_TYPES = ["1:4", "1:6"]

WEIGHT_RANGES = {
    "M10": (8.100, 8.300), "M15": (8.100, 8.300), "M20": (8.100, 8.300),
    "M25": (8.180, 8.350), "M30": (8.100, 8.350), "M35": (8.100, 8.350),
    "M40": (8.100, 8.350), "M45": (8.200, 8.400),
    "1:4": (0.800, 0.835), "1:6": (0.800, 0.835),
}

STRENGTH_7D_RANGES = {
    "M10": (214.00, 267.40), "M15": (290.10, 320.50), "M20": (366.10, 410.10),
    "M25": (442.10, 490.10), "M30": (518.10, 560.10), "M35": (595.10, 632.80),
    "M40": (669.10, 728.10), "M45": (735.10, 788.10),
    "1:4": (25.20, 33.90),  "1:6": (15.20, 25.00),
}

STRENGTH_28D_RANGES = {
    "M10": (320.10, 365.50), "M15": (433.10, 480.10), "M20": (547.10, 590.10),
    "M25": (660.10, 710.10), "M30": (770.10, 812.10), "M35": (880.90, 925.10),
    "M40": (995.10, 1038.10), "M45": (1105.35, 1150.10),
    "1:4": (40.60, 50.10),   "1:6": (25.20, 33.90),
}

ALL_TYPES = CONCRETE_GRADES + MORTAR_TYPES


def _generate_unique_values(min_val, max_val, count, decimals=2, min_gap=0.0):
    """Generate *count* unique random values in [min_val, max_val] with a minimum gap."""
    values = []
    max_attempts = 2000

    attempts = 0
    while len(values) < count and attempts < max_attempts:
        val = round(np.random.uniform(min_val, max_val), decimals)
        if val not in values and all(abs(val - v) >= min_gap for v in values):
            values.append(val)
        attempts += 1

    # Fallback – nudge duplicates
    while len(values) < count:
        val = values[-1] + round(np.random.uniform(0.001, 0.009), 3)
        if min_val <= val <= max_val and val not in values:
            values.append(round(val, decimals))

    return values[:count]


def generate_row(grade_or_type):
    """
    Generate a single row of test data for the given grade / mortar type.

    Returns
    -------
    weights : list[float]   — 6 values
    strength_7d : list[float] — 3 values
    strength_28d : list[float] — 3 values
    """
    w_min, w_max = WEIGHT_RANGES[grade_or_type]
    s7_min, s7_max = STRENGTH_7D_RANGES[grade_or_type]
    s28_min, s28_max = STRENGTH_28D_RANGES[grade_or_type]

    is_mortar = grade_or_type in MORTAR_TYPES
    weight_decimals = 3
    weight_gap = 0.005 if is_mortar else 0.015
    strength_gap = 1.0 if is_mortar else 5.0

    weights = _generate_unique_values(w_min, w_max, 6, decimals=weight_decimals, min_gap=weight_gap)
    np.random.shuffle(weights)

    strength_7d = _generate_unique_values(s7_min, s7_max, 3, decimals=2, min_gap=strength_gap)
    np.random.shuffle(strength_7d)

    strength_28d = _generate_unique_values(s28_min, s28_max, 3, decimals=2, min_gap=strength_gap)
    np.random.shuffle(strength_28d)

    return weights, strength_7d, strength_28d


def generate_rows(grade_or_type, count):
    """Yield *count* rows of (weights, strength_7d, strength_28d)."""
    for _ in range(count):
        yield generate_row(grade_or_type)


def grade_display_name(grade_or_type):
    """Friendly display name for a grade/type."""
    if grade_or_type in MORTAR_TYPES:
        return f"Mortar {grade_or_type}"
    return grade_or_type
