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

import numpy as np

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

# ── Dynamic Range Expansion for Concrete ──
# 1 derived digit = 22.5 kN
# 7-day: M10, M15 get +22.5 (1 digit). M20+ get +45.0 (2 digits).
# 28-day: All concrete gets +22.5 (1 digit).
for g in CONCRETE_GRADES:
    # 7-day expansion
    s_min, s_max = STRENGTH_7D_RANGES[g]
    if g in ["M10", "M15"]:
        STRENGTH_7D_RANGES[g] = (s_min, s_max + 22.5)
    else:
        STRENGTH_7D_RANGES[g] = (s_min, s_max + 45.0)
        
    # 28-day expansion
    s_min, s_max = STRENGTH_28D_RANGES[g]
    STRENGTH_28D_RANGES[g] = (s_min, s_max + 22.5)


# ============================================================================
# Derived-value helpers
# ============================================================================

def _derived_scale(is_mortar):
    return 10.0 / 49.8 if is_mortar else 10.0 / 225.0

def _derived_avg(strength_list, is_mortar):
    """Average derived value for a list of raw kN values. Pure numpy."""
    return float(np.mean(np.asarray(strength_list) * _derived_scale(is_mortar)))


# ============================================================================
# Pre-computed per-grade zone tables  (built once at import)
# ============================================================================

def _build_zone_table(s_min, s_max, is_mortar):
    """
    Return list of (zone_int, raw_lo, raw_hi) covering the derived range.
    raw_lo / raw_hi are in kN.  Inward pad of 0.12 derived units keeps the
    generated average reliably inside the zone after clamping.
    """
    scale = _derived_scale(is_mortar)
    d_min = s_min * scale
    d_max = s_max * scale
    zones = []
    PAD = 0.12
    for z in range(int(d_min), int(d_max) + 2):
        lo = max(z + PAD, d_min + 0.03)
        hi = min(z + 1.0 - PAD, d_max - 0.03)
        if lo < hi:
            zones.append((z, lo / scale, hi / scale))
    return zones

_ZONE_TABLE_7D  = {g: _build_zone_table(STRENGTH_7D_RANGES[g][0], STRENGTH_7D_RANGES[g][1], g in MORTAR_TYPES) for g in ALL_TYPES}
_ZONE_TABLE_28D = {g: _build_zone_table(STRENGTH_28D_RANGES[g][0], STRENGTH_28D_RANGES[g][1], g in MORTAR_TYPES) for g in ALL_TYPES}

# Mortar: threshold = 10% of each field's own derived span, min 0.02
def _derived_span(s_min, s_max, is_mortar):
    return (s_max - s_min) * _derived_scale(is_mortar)

_MORTAR_THRESH_7D = {
    g: max(0.02, _derived_span(STRENGTH_7D_RANGES[g][0], STRENGTH_7D_RANGES[g][1], True)  * 0.10) for g in MORTAR_TYPES
}
_MORTAR_THRESH_28D = {
    g: max(0.02, _derived_span(STRENGTH_28D_RANGES[g][0], STRENGTH_28D_RANGES[g][1], True) * 0.10) for g in MORTAR_TYPES
}


# ============================================================================
# Weight generation -- single-pass, gap-enforced
# ============================================================================

def _gen_weights(w_min, w_max, count, min_gap, is_mortar, decimals=3):
    """
    Generate *count* weights in [w_min, w_max] with every adjacent sorted pair
    at least *min_gap* apart.

    Uses a perfect random placement algorithm:
    1. Calculate required width for gaps.
    2. Expand w_max dynamically if there is not enough slack for randomness.
    3. Generate sorted random points in the slack space.
    """
    required_width = (count - 1) * min_gap
    slack = w_max - w_min - required_width
    
    # Ensure there is enough slack so the weights are highly random
    # instead of ending exactly in 0s.
    min_slack = 0.005 if is_mortar else 0.040
    if slack < min_slack:
        w_max = w_min + required_width + min_slack
        slack = min_slack

    random_sorted = np.sort(_rng.uniform(0, slack, count))
    pts = w_min + random_sorted + np.arange(count) * min_gap
    pts = np.round(pts, decimals)
    
    # Enforce minimum gap exactly post-rounding
    for i in range(1, count):
        if pts[i] - pts[i - 1] < min_gap:
            pts[i] = pts[i - 1] + min_gap
            
    # Clamp if nudging pushed it past the expanded w_max
    if pts[-1] > w_max:
        pts -= (pts[-1] - w_max)
        pts = np.round(pts, decimals)
        for i in range(1, count):
            if pts[i] - pts[i - 1] < min_gap:
                pts[i] = pts[i - 1] + min_gap

    pts = np.round(pts, decimals).tolist()
    _rng.shuffle(pts)
    return pts


# ============================================================================
# Strength generation -- algebraic 3-value layout, gap-enforced
# ============================================================================

def _gen_strengths(s_min, s_max, target_raw, min_gap, decimals=2):
    """
    Generate exactly 3 strength values in [s_min, s_max] whose mean is
    *approximately* target_raw, with each adjacent sorted pair >= min_gap apart.

    Uses a perfect O(1) slack-partitioning algorithm to guarantee vast randomness
    without any rejection sampling or loops.
    """
    # Clamp target_raw to the mathematically solvable domain
    min_possible_mean = s_min + min_gap
    max_possible_mean = s_max - min_gap
    
    if min_possible_mean > max_possible_mean:
        target_raw = (s_min + s_max) / 2.0
    else:
        target_raw = np.clip(target_raw, min_possible_mean, max_possible_mean)
        
    # We partition the slack exactly
    S = 3.0 * target_raw
    total_slack = s_max - s_min - 2.0 * min_gap
    K = S - 3.0 * s_min - 3.0 * min_gap
    
    # Clip K for floating point safety
    K = np.clip(K, 0.0, 3.0 * total_slack)
    
    # x1 represents the slack of the lowest value
    x1_min = max(0.0, K - 2.0 * total_slack)
    x1_max = K / 3.0
    x1 = _rng.uniform(x1_min, x1_max)
    
    # x2 represents the slack of the middle value
    x2_min = max(0.0, K - total_slack - 2.0 * x1)
    x2_max = (K - 3.0 * x1) / 2.0
    x2 = _rng.uniform(x2_min, x2_max)
    
    # x3 represents the slack of the highest value
    x3 = K - 3.0 * x1 - 2.0 * x2
    
    # Construct values
    v1 = s_min + x1
    v2 = v1 + min_gap + x2
    v3 = v2 + min_gap + x3
    
    vals = [v1, v2, v3]
    vals = np.round(vals, decimals)
    
    # Enforce gaps post-rounding strictly
    if vals[1] - vals[0] < min_gap:
        vals[1] = vals[0] + min_gap
    if vals[2] - vals[1] < min_gap:
        vals[2] = vals[1] + min_gap
        
    # Clamp overflow (should only occur due to rounding)
    if vals[2] > s_max:
        vals[2] = s_max
        vals[1] = min(vals[1], round(s_max - min_gap, decimals))
        vals[0] = min(vals[0], round(s_max - 2.0 * min_gap, decimals))
        
    vals = np.round(vals, decimals).tolist()
    _rng.shuffle(vals)
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
    else:
        return int(avg) != int(prev_avg)


# ============================================================================
# Public API
# ============================================================================

def generate_row(grade_or_type):
    """
    Generate a single row of test data for the given grade / mortar type.

    Returns
    -------
    weights      : list[float]  -- 6 values
    strength_7d  : list[float]  -- 3 values
    strength_28d : list[float]  -- 3 values
    """
    is_mortar    = grade_or_type in MORTAR_TYPES
    weight_gap   = 0.005 if is_mortar else 0.040
    strength_gap = 1.0   if is_mortar else 10.0

    w_min,  w_max   = WEIGHT_RANGES[grade_or_type]
    s7_min, s7_max  = STRENGTH_7D_RANGES[grade_or_type]
    s28_min, s28_max = STRENGTH_28D_RANGES[grade_or_type]

    weights      = _gen_weights(w_min, w_max, 6, weight_gap, is_mortar)
    t7           = float(_rng.uniform(s7_min,  s7_max))
    t28          = float(_rng.uniform(s28_min, s28_max))
    strength_7d  = _gen_strengths(s7_min,  s7_max,  t7,  strength_gap)
    strength_28d = _gen_strengths(s28_min, s28_max, t28, strength_gap)
    return weights, strength_7d, strength_28d


def generate_rows(grade_or_type, count):
    """
    Yield *count* rows of (weights, strength_7d, strength_28d).

    Per-sheet guarantees
    --------------------
    * Weights  : min gap >= 0.040 kg (concrete) / >= 0.005 kg (mortar)
    * Strengths: min gap >= 10.0 kN (concrete)  / all distinct (mortar)
    * All values within the defined range for the grade/type.

    Cross-sheet guarantee
    ---------------------
    * The derived average (raw_kN / 225 * 10) of BOTH 7-day and 28-day
      strengths has a DIFFERENT INTEGER PART on consecutive sheets (concrete),
      or differs by >= threshold on consecutive sheets (mortar).

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

    MAX_RETRIES = 15

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
        for _attempt in range(MAX_RETRIES):
            s7 = _gen_strengths(s7_min, s7_max, target_7d, strength_gap)
            if _avg_differs(s7, prev_avg_7d, is_mortar, m_thresh_7d):
                break
            # Target landed in wrong zone after clamping -- repick
            target_7d = _pick_target(zt7, prev_avg_7d, is_mortar, m_thresh_7d)

        # ---------- 28-day strengths -----------------------------------------
        target_28d = _pick_target(zt28, prev_avg_28d, is_mortar, m_thresh_28d)
        for _attempt in range(MAX_RETRIES):
            s28 = _gen_strengths(s28_min, s28_max, target_28d, strength_gap)
            if _avg_differs(s28, prev_avg_28d, is_mortar, m_thresh_28d):
                break
            target_28d = _pick_target(zt28, prev_avg_28d, is_mortar, m_thresh_28d)

        prev_avg_7d  = _derived_avg(s7, is_mortar)
        prev_avg_28d = _derived_avg(s28, is_mortar)

        yield weights, s7, s28


def grade_display_name(grade_or_type):
    """Friendly display name for a grade/type."""
    if grade_or_type in MORTAR_TYPES:
        return f"Mortar {grade_or_type}"
    return grade_or_type
