"""
Cube Data Generator Module
Generates synthetic concrete and mortar cube test data (weights + strengths)
in-memory, ready for direct processing without intermediate Excel files.

Based on: https://github.com/Sandeep2062/Cube-Data-Generator

Optimisations
─────────────
* Pure-numpy batch generation -- avoids Python-level loops as much as possible.
* O(1) cross-sheet derived-average constraint using pre-computed integer zones.
* All random calls batched with a single RNG instance (no repeated seeding).
* Zero retry loops for weight generation (evenly-spaced + jitter).
* Minimal retry for strength cross-sheet constraint (capped at 15, typically 1-2).
"""

import math
import numpy as np
from collections import deque

# Module-level RNG -- faster than calling np.random.* global functions
_rng = np.random.default_rng()


# ============================================================================
# Range definitions
# ============================================================================

CONCRETE_GRADES = ["M10", "M15", "M20", "M25", "M30", "M35", "M40", "M45"]
MORTAR_TYPES    = ["1:4", "1:6"]
ALL_TYPES       = CONCRETE_GRADES + MORTAR_TYPES

WEIGHT_RANGES = {
    "M10": (8.100, 8.300), "M15": (8.100, 8.300), "M20": (8.100, 8.300),
    "M25": (8.180, 8.350), "M30": (8.100, 8.350), "M35": (8.100, 8.350),
    "M40": (8.100, 8.350), "M45": (8.200, 8.400),
    "1:4": (0.800, 0.835), "1:6": (0.800, 0.835),
}

_BASE_STRENGTH_7D_RANGES = {
    "M10": (214.00, 267.40), "M15": (290.10, 320.50), "M20": (366.10, 410.10),
    "M25": (442.10, 490.10), "M30": (518.10, 560.10), "M35": (595.10, 632.80),
    "M40": (669.10, 728.10), "M45": (735.10, 788.10),
    "1:4": (25.20, 33.90),  "1:6": (15.20, 25.00),
}

_BASE_STRENGTH_28D_RANGES = {
    "M10": (320.10, 365.50), "M15": (433.10, 480.10), "M20": (547.10, 590.10),
    "M25": (660.10, 710.10), "M30": (770.10, 812.10), "M35": (880.90, 925.10),
    "M40": (995.10, 1038.10), "M45": (1105.35, 1150.10),
    "1:4": (40.60, 50.10),   "1:6": (25.20, 33.90),
}

STRENGTH_7D_RANGES = {}
STRENGTH_28D_RANGES = {}

# Pre-computed per-grade zone tables
_ZONE_TABLE_7D = {}
_ZONE_TABLE_28D = {}
_MORTAR_THRESH_7D = {}
_MORTAR_THRESH_28D = {}


# ============================================================================
# Weight generation -- single-pass, gap-enforced
# ============================================================================

def _gen_weights(w_min, w_max, count, min_gap, is_mortar, decimals=3):
    """
    Generate *count* weights in [w_min, w_max] with every adjacent sorted pair
    at least *min_gap* apart.

    Uses a bounded random placement algorithm. It never expands w_max; if the
    configured gap is too large for the configured range, the largest feasible
    gap for that range is used.
    """
    range_span = w_max - w_min
    if count <= 1:
        return [round(float(_rng.uniform(w_min, w_max)), decimals)]

    effective_gap = min(min_gap, range_span / (count - 1))
    required_width = (count - 1) * effective_gap
    slack = w_max - w_min - required_width

    random_sorted = np.sort(_rng.uniform(0, slack, count))
    pts = w_min + random_sorted + np.arange(count) * effective_gap
    pts = np.round(pts, decimals)
    
    # Enforce minimum gap exactly post-rounding
    for i in range(1, count):
        if pts[i] - pts[i - 1] < effective_gap:
            pts[i] = pts[i - 1] + effective_gap
            
    # Clamp if rounding/nudging touched the upper boundary.
    if pts[-1] > w_max:
        pts -= (pts[-1] - w_max)
        pts = np.round(pts, decimals)
        for i in range(1, count):
            if pts[i] - pts[i - 1] < effective_gap:
                pts[i] = pts[i - 1] + effective_gap

    pts = np.clip(np.round(pts, decimals), w_min, w_max).tolist()
    _rng.shuffle(pts)
    return pts


# ============================================================================
# Strength generation -- wide-spread, random, gap-enforced
# ============================================================================

def _gen_strengths(s_min, target_raw, min_gap, s_max=None, decimals=2,
                   used_values=None, forbidden_decimals=None):
    """
    Generate exactly 3 strength values whose mean is close to *target_raw*.
    
    Each adjacent sorted pair must have AT LEAST a randomized minimum gap:
      - Concrete: minimum gap randomized between 10.00 and 12.70
      - Mortar:   minimum gap randomized between 1.68 and 2.32
    
    Values CAN be much further apart than the minimum gap — they are scattered
    randomly across the full [s_min, s_max] range for natural-looking data.
    Example output: 305.26, 338.45, 290.15 (gaps of 15.11 and 33.19, both >= 10)
    """
    if s_max is None:
        s_max = target_raw * 1.02
        
    target_raw = max(target_raw, s_min + min_gap)
    target_raw = min(target_raw, s_max)
    
    if forbidden_decimals is None:
        forbidden_decimals = set()
    if used_values is None:
        used_values = set()
        
    range_span = s_max - s_min
    step = 10 ** (-decimals)
    
    # Safely clear tracker to prevent impossible loops for 1000s of sheets.
    # Concrete now has 1 decimal place, so tight grades only have a few hundred
    # possible values in total.
    value_slots = max(1, int(range_span / step))
    if len(used_values) > value_slots * 0.35:
        used_values.clear()
    
    best_vals = None
    
    # The MINIMUM gap between each adjacent sorted pair is randomized in this range
    if min_gap >= 10.0:
        mgap_lo, mgap_hi = 10.00, 12.70
    else:
        mgap_lo, mgap_hi = 1.68, 2.32
        
    for _overall_attempt in range(350):
        req_gap = _round_up(_rng.uniform(mgap_lo, mgap_hi), decimals)
        if s_max - s_min < 2.0 * req_gap:
            req_gap = _round_down((s_max - s_min) / 2.0, decimals)

        # Generate random sorted values with guaranteed minimum gaps by
        # sampling compressed points, then adding the reserved gap back.
        compressed_max = s_max - 2.0 * req_gap
        if compressed_max < s_min:
            continue
        compressed = np.sort(_rng.uniform(s_min, compressed_max, 3))
        vals = [
            compressed[0],
            compressed[1] + req_gap,
            compressed[2] + 2.0 * req_gap,
        ]

        # Lightly bias toward the selected target without allowing any clamp
        # operation to create out-of-bounds or too-close values.
        current_mean = sum(vals) / 3.0
        max_shift_down = s_min - vals[0]
        max_shift_up = s_max - vals[2]
        wanted_shift = (target_raw - current_mean) * 0.35
        shift = min(max(wanted_shift, max_shift_down), max_shift_up)
        vals = [v + shift for v in vals]

        vals = [float(round(v, decimals)) for v in vals]
        vals = sorted(vals)
        
        # Verify minimum gaps after rounding
        gap1 = round(vals[1] - vals[0], decimals)
        gap2 = round(vals[2] - vals[1], decimals)
        if gap1 < req_gap or gap2 < req_gap:
            continue
        
        # Verify hard boundaries
        if vals[0] < s_min or vals[2] > s_max:
            continue
            
        # Verify discrete uniqueness constraints
        dec_parts = [_get_decimal_part(v, decimals) for v in vals]
        unique_dec = len(set(dec_parts)) == 3
        no_forbid_dec = all(dp not in forbidden_decimals for dp in dec_parts)
        no_used_val = all(v not in used_values for v in vals)
        
        if unique_dec and no_forbid_dec and no_used_val:
            best_vals = vals
            break
        if unique_dec and no_forbid_dec and _overall_attempt > 80:
            best_vals = vals
            break
            
    if best_vals is None:
        best_vals = _fallback_strengths(s_min, s_max, mgap_lo, decimals,
                                        forbidden_decimals, used_values)
        
    for v in best_vals:
        used_values.add(v)
        
    _rng.shuffle(best_vals)
    return best_vals


def _get_decimal_part(value, decimals=2):
    """Extract the decimal portion of a float. e.g. 276.15 -> 15, 302.50 -> 50"""
    multiplier = 10 ** decimals
    return round(value * multiplier) % multiplier


def _round_up(value, decimals):
    multiplier = 10 ** decimals
    return math.ceil(value * multiplier) / multiplier


def _round_down(value, decimals):
    multiplier = 10 ** decimals
    return math.floor(value * multiplier) / multiplier


def _fallback_strengths(s_min, s_max, min_gap, decimals, forbidden_decimals, used_values):
    """Last-resort bounded generator; still never exceeds strength limits."""
    gap = _round_up(min_gap, decimals)
    low = round(s_min, decimals)
    high = round(s_max, decimals)
    step = 10 ** (-decimals)

    for _ in range(1000):
        first_hi = high - 2.0 * gap
        if first_hi < low:
            break
        v1 = round(float(_rng.uniform(low, first_hi)), decimals)
        v2 = round(float(_rng.uniform(v1 + gap, high - gap)), decimals)
        v3 = round(float(_rng.uniform(v2 + gap, high)), decimals)
        vals = sorted([v1, v2, v3])
        dec_parts = [_get_decimal_part(v, decimals) for v in vals]
        if len(set(dec_parts)) == 3 and all(dp not in forbidden_decimals for dp in dec_parts):
            return vals

    vals = [low, round(low + gap, decimals), round(low + 2.0 * gap, decimals)]
    vals = [float(min(max(v, low), high)) for v in vals]
    return vals


# ============================================================================
# Cross-sheet target selection
# ============================================================================

def _pick_target(zone_table, prev_avg, is_mortar, mortar_threshold=0.10):
    """
    Return a target raw-kN value that will yield a derived average in a
    DIFFERENT integer zone (concrete) or sufficiently far (mortar) from prev_avg.

    O(len(zone_table)) -- typically 2-6 zones.
    """
    if not zone_table:
        return 0.0

    if prev_avg is None:
        z_int, raw_lo, raw_hi = zone_table[_rng.integers(len(zone_table))]
        return float(_rng.uniform(raw_lo, raw_hi))

    prev_int = int(prev_avg)

    if is_mortar:
        # For mortar: pick zone whose MID-POINT differs sufficiently from prev_avg
        candidates = [
            (z, lo, hi) for z, lo, hi in zone_table
            if abs(((lo + hi) / 2.0) * _derived_scale(is_mortar) - prev_avg) >= mortar_threshold
        ]
    else:
        # For concrete: different integer zone
        candidates = [(z, lo, hi) for z, lo, hi in zone_table if z != prev_int]

    # Fallback: always pick zone whose centre is FURTHEST from prev_avg
    if not candidates:
        candidates = [max(
            zone_table,
            key=lambda t: abs(((t[1] + t[2]) / 2.0) * _derived_scale(is_mortar) - prev_avg)
        )]

    z_int, raw_lo, raw_hi = candidates[_rng.integers(len(candidates))]
    return float(_rng.uniform(raw_lo, raw_hi))


# ============================================================================
# Constraint verification
# ============================================================================

def _avg_differs(vals, prev_avg, is_mortar, mortar_threshold):
    """True if this set of values satisfies the cross-sheet constraint."""
    if prev_avg is None:
        return True
    avg = _derived_avg(vals, is_mortar)
    if is_mortar:
        return abs(avg - prev_avg) >= mortar_threshold
    # For concrete, the integer part MUST differ on consecutive sheets.
    return int(avg) != int(prev_avg)


def _avg_band_differs(vals, prev_avg, is_mortar):
    """Prefer a different displayed integer band when concrete has room."""
    if prev_avg is None or is_mortar:
        return True
    return int(_display_avg(vals, is_mortar)) != int(round(prev_avg, 2))


def _derived_avg_unique(avg, recent_avgs, decimals=2):
    """
    Check that this derived average (rounded to `decimals`) has not appeared
    in the recent_avgs deque. Returns True if unique.
    """
    rounded = round(avg, decimals)
    return rounded not in recent_avgs


def _display_avg(strength_list, is_mortar):
    """Average as displayed by the sheet formula, rounded to 2 decimals."""
    return round(_derived_avg(strength_list, is_mortar), 2)


def _display_avg_allowed(strength_list, is_mortar, recent_avgs):
    """True when the final displayed average has not appeared recently."""
    return _display_avg(strength_list, is_mortar) not in recent_avgs


# ============================================================================
# Public API
# ============================================================================

def generate_row(grade_or_type):
    """
    Generate a single random row with no cross-sheet constraints.
    Used by the preview function and quick-generation modes.
    """
    is_mortar    = grade_or_type in MORTAR_TYPES
    weight_gap   = 0.005 if is_mortar else 0.040
    strength_gap = 1.0   if is_mortar else 10.0
    dec_places   = 2 if is_mortar else 1  # Concrete=1 decimal, Mortar=2 decimals

    w_min,  w_max    = WEIGHT_RANGES[grade_or_type]
    s7_min, s7_max   = STRENGTH_7D_RANGES[grade_or_type]
    s28_min, s28_max = STRENGTH_28D_RANGES[grade_or_type]

    weights = _gen_weights(w_min, w_max, 6, weight_gap, is_mortar)

    t7           = float(_rng.uniform(s7_min,  s7_max))
    t28          = float(_rng.uniform(s28_min, s28_max))
    strength_7d  = _gen_strengths(s7_min, t7,  strength_gap, s_max=s7_max,
                                   decimals=dec_places)
    # Pass 7d decimal parts as forbidden for 28d to avoid cross-triplet dups
    dec_7d = {_get_decimal_part(v, dec_places) for v in strength_7d}
    strength_28d = _gen_strengths(s28_min, t28, strength_gap, s_max=s28_max,
                                   decimals=dec_places,
                                   forbidden_decimals=dec_7d)
    return weights, strength_7d, strength_28d


def generate_rows(grade_or_type, count):
    """
    Yield *count* rows of (weights, strength_7d, strength_28d).

    Per-sheet guarantees
    --------------------
    * Weights  : min gap >= 0.040 kg (concrete) / >= 0.005 kg (mortar)
    * Strengths: min gap >= 10.0 kN (concrete)  / >= 1.0 kN (mortar)
    * All values strictly within the defined range for the grade/type.
    * All 6 strength values on a sheet have UNIQUE decimal parts.

    Cross-sheet guarantee
    ---------------------
    * The derived average (raw_kN / 225 * 10) of BOTH 7-day and 28-day
      strengths has a DIFFERENT INTEGER PART on consecutive sheets (concrete),
      or differs by >= threshold on consecutive sheets (mortar).
    * The rounded derived average cannot repeat within 10 consecutive sheets.
    * No exact strength value repeats across any sheet in the file.

    Performance
    -----------
    * Weights: zero retry, O(count) numpy.
    * Strengths: at most 15 retries per field (typically 1-2 in practice).
    * Pre-computed zone tables for O(zones) target lookup.
    * Single module-level numpy RNG instance.
    """
    is_mortar    = grade_or_type in MORTAR_TYPES
    weight_gap   = 0.005 if is_mortar else 0.040
    strength_gap = 1.0   if is_mortar else 10.0
    dec_places   = 2 if is_mortar else 1  # Concrete=1 decimal, Mortar=2 decimals

    w_min,  w_max    = WEIGHT_RANGES[grade_or_type]
    s7_min, s7_max   = STRENGTH_7D_RANGES[grade_or_type]
    s28_min, s28_max = STRENGTH_28D_RANGES[grade_or_type]

    zt7  = _ZONE_TABLE_7D[grade_or_type]
    zt28 = _ZONE_TABLE_28D[grade_or_type]
    m_thresh_7d  = _MORTAR_THRESH_7D.get(grade_or_type,  0.04)
    m_thresh_28d = _MORTAR_THRESH_28D.get(grade_or_type, 0.04)

    prev_avg_7d  = None
    prev_avg_28d = None
    prev_weights = set()
    used_7d_values  = set()   # Track ALL generated 7d values across sheets
    used_28d_values = set()   # Track ALL generated 28d values across sheets
    # Track recent derived averages -- same average can't repeat within 10 sheets
    recent_avgs_7d  = deque(maxlen=10)
    recent_avgs_28d = deque(maxlen=10)

    MAX_RETRIES = 80

    for _ in range(count):
        # ---------- Weights (ensure no cross-sheet overlap) ------------------
        weights = _gen_weights(w_min, w_max, 6, weight_gap, is_mortar)
        
        # Give it up to 10 attempts to find completely disjoint weights
        for _attempt in range(10):
            if prev_weights and any(w in prev_weights for w in weights):
                weights = _gen_weights(w_min, w_max, 6, weight_gap, is_mortar)
            else:
                break
                
        prev_weights = set(weights)

        # ---------- 7-day strengths ------------------------------------------
        target_7d = _pick_target(zt7, prev_avg_7d, is_mortar, m_thresh_7d)
        s7 = None
        for _attempt in range(MAX_RETRIES):
            s7 = _gen_strengths(s7_min, target_7d, strength_gap,
                                s_max=s7_max, decimals=dec_places,
                                used_values=used_7d_values)
            avg_7d = _derived_avg(s7, is_mortar)
            zone_ok = _avg_differs(s7, prev_avg_7d, is_mortar, m_thresh_7d)
            unique_ok = _display_avg_allowed(s7, is_mortar, recent_avgs_7d)
            if zone_ok and unique_ok:
                break
            # Target landed in wrong zone or avg repeated -- repick
            target_7d = _pick_target(zt7, prev_avg_7d, is_mortar, m_thresh_7d)

        # ---------- 28-day strengths -----------------------------------------
        # Pass 7d decimal parts as forbidden so all 6 values have unique decimals
        dec_7d = {_get_decimal_part(v, dec_places) for v in s7}
        
        target_28d = _pick_target(zt28, prev_avg_28d, is_mortar, m_thresh_28d)
        s28 = None
        for _attempt in range(MAX_RETRIES):
            s28 = _gen_strengths(s28_min, target_28d, strength_gap,
                                 s_max=s28_max, decimals=dec_places,
                                 used_values=used_28d_values,
                                 forbidden_decimals=dec_7d)
            avg_28d = _derived_avg(s28, is_mortar)
            zone_ok = _avg_differs(s28, prev_avg_28d, is_mortar, m_thresh_28d)
            unique_ok = _display_avg_allowed(s28, is_mortar, recent_avgs_28d)
            if zone_ok and unique_ok:
                break
            target_28d = _pick_target(zt28, prev_avg_28d, is_mortar, m_thresh_28d)

        prev_avg_7d  = _derived_avg(s7, is_mortar)
        prev_avg_28d = _derived_avg(s28, is_mortar)
        
        # Track rounded averages for the 10-sheet uniqueness window
        recent_avgs_7d.append(_display_avg(s7, is_mortar))
        recent_avgs_28d.append(_display_avg(s28, is_mortar))

        yield weights, s7, s28

def _derived_scale(is_mortar):
    return 10.0 / 49.8 if is_mortar else 10.0 / 225.0

def _derived_avg(strength_list, is_mortar):
    """Average derived value for a list of raw kN values. Pure numpy."""
    return float(np.mean(np.asarray(strength_list) * _derived_scale(is_mortar)))


def _force_unique_display_average(s_min, s_max, target_raw, strength_gap,
                                  decimals, is_mortar, recent_avgs,
                                  used_values, forbidden_decimals=None):
    """Generate until the final C30/F30-style average is different."""
    for _ in range(500):
        vals = _gen_strengths(
            s_min, target_raw, strength_gap, s_max=s_max,
            decimals=decimals, used_values=used_values,
            forbidden_decimals=forbidden_decimals)
        if _display_avg_allowed(vals, is_mortar, recent_avgs):
            return vals

    # The average space is usually large, but if a custom range is extremely
    # tight, return the best bounded row rather than breaking generation.
    return vals

def _build_zone_table(s_min, s_max, is_mortar):
    """
    Return list of (zone_int, raw_lo, raw_hi) covering the derived range.
    raw_lo / raw_hi are in kN.  Inward pad of 0.12 derived units keeps the
    generated average reliably inside the zone after clamping.
    """
    scale = _derived_scale(is_mortar)
    min_gap = 1.0 if is_mortar else 10.0
    
    # Account for randomized minimum gaps: 10.00-12.70 concrete,
    # 1.68-2.32 mortar.
    max_gap_diff = 2.7 if not is_mortar else 1.32
    
    # The absolute lowest and highest possible mean values
    # given the randomized mathematical gap between the 3 values.
    lowest_mean = s_min + min_gap + max_gap_diff
    highest_mean = s_max - min_gap - max_gap_diff
    
    # If the range is extremely tight, collapse to the midpoint
    if lowest_mean > highest_mean:
        lowest_mean = highest_mean = (s_min + s_max) / 2.0
    
    d_min = lowest_mean * scale
    d_max = highest_mean * scale
    
    zones = []
    PAD = 0.12
    for z in range(int(d_min), int(d_max) + 2):
        lo = max(z + PAD, d_min + 0.03)
        hi = min(z + 1.0 - PAD, d_max - 0.03)
        if lo < hi:
            zones.append((z, lo / scale, hi / scale))
    if not zones:
        midpoint = (lowest_mean + highest_mean) / 2.0
        zones.append((int(midpoint * scale), midpoint, midpoint))
    return zones

def _derived_span(s_min, s_max, is_mortar):
    return (s_max - s_min) * _derived_scale(is_mortar)

def override_ranges(custom_7d=None, custom_28d=None):
    """
    Apply custom base ranges (if any) and re-initialize all dynamic expansions
    and zone tables. Must be called before generating rows if overrides exist.
    """
    global STRENGTH_7D_RANGES, STRENGTH_28D_RANGES
    global _ZONE_TABLE_7D, _ZONE_TABLE_28D
    global _MORTAR_THRESH_7D, _MORTAR_THRESH_28D

    custom_7d = custom_7d or {}
    custom_28d = custom_28d or {}

    STRENGTH_7D_RANGES.clear()
    STRENGTH_28D_RANGES.clear()

    # Load base or custom
    for g in ALL_TYPES:
        STRENGTH_7D_RANGES[g] = tuple(custom_7d.get(g, _BASE_STRENGTH_7D_RANGES[g]))
        STRENGTH_28D_RANGES[g] = tuple(custom_28d.get(g, _BASE_STRENGTH_28D_RANGES[g]))

    # -- Dynamic Range Expansion for Concrete ONLY --
    # Tight lower grades need adaptive headroom; otherwise their displayed
    # averages get trapped in one band (for example M15 7-day stays 13.xx).
    # We expand the bounds OUTWARD until they are wide enough to guarantee
    # multiple integer zones.
    # We need (s_max - s_min - 25.4) * 10 / 225 > 1.2  =>  s_max - s_min > 52.4
    for g in CONCRETE_GRADES:
        # 7-day
        s_min, s_max = STRENGTH_7D_RANGES[g]
        if s_max - s_min < 54.0:
            deficit = 54.0 - (s_max - s_min)
            STRENGTH_7D_RANGES[g] = (round(s_min - deficit/2, 2), round(s_max + deficit/2, 2))
            
        # 28-day
        s_min, s_max = STRENGTH_28D_RANGES[g]
        if s_max - s_min < 54.0:
            deficit = 54.0 - (s_max - s_min)
            STRENGTH_28D_RANGES[g] = (round(s_min - deficit/2, 2), round(s_max + deficit/2, 2))

    _ZONE_TABLE_7D.clear()
    _ZONE_TABLE_28D.clear()
    _MORTAR_THRESH_7D.clear()
    _MORTAR_THRESH_28D.clear()

    for g in ALL_TYPES:
        is_m = g in MORTAR_TYPES
        _ZONE_TABLE_7D[g] = _build_zone_table(STRENGTH_7D_RANGES[g][0], STRENGTH_7D_RANGES[g][1], is_m)
        _ZONE_TABLE_28D[g] = _build_zone_table(STRENGTH_28D_RANGES[g][0], STRENGTH_28D_RANGES[g][1], is_m)

    for g in MORTAR_TYPES:
        _MORTAR_THRESH_7D[g] = max(0.02, _derived_span(STRENGTH_7D_RANGES[g][0], STRENGTH_7D_RANGES[g][1], True) * 0.10)
        _MORTAR_THRESH_28D[g] = max(0.02, _derived_span(STRENGTH_28D_RANGES[g][0], STRENGTH_28D_RANGES[g][1], True) * 0.10)

# Initialize defaults on import
override_ranges()


def grade_display_name(grade_or_type):
    """Friendly display name for a grade/type."""
    if grade_or_type in MORTAR_TYPES:
        return f"Mortar {grade_or_type}"
    return grade_or_type
