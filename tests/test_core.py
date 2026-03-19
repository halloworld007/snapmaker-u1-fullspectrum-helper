"""
Unit tests for u1 FullSpectrum core functions.
Run with: pytest tests/test_core.py -v
"""
import sys
import os
import math
import pytest

# ── Inline the pure functions so we don't need PySide6 to run tests ──────────

def hex_to_rgb(hex_str):
    s = str(hex_str).strip().lstrip("#")
    if len(s) == 8: s = s[:6]
    if len(s) != 6:
        return (128, 128, 128)
    try:
        return tuple(int(s[i:i+2], 16) for i in (0, 2, 4))
    except ValueError:
        return (128, 128, 128)

def rgb_to_hex(r, g, b):
    return "#{:02X}{:02X}{:02X}".format(int(max(0, min(255, r))),
                                         int(max(0, min(255, g))),
                                         int(max(0, min(255, b))))

_LAB_CACHE = {}

def rgb_to_lab(rgb):
    key = (int(rgb[0]), int(rgb[1]), int(rgb[2]))
    if key in _LAB_CACHE:
        return _LAB_CACHE[key]
    r, g, b = [x / 255.0 for x in key]
    r = (r / 12.92) if r <= 0.04045 else ((r + 0.055) / 1.055) ** 2.4
    g = (g / 12.92) if g <= 0.04045 else ((g + 0.055) / 1.055) ** 2.4
    b = (b / 12.92) if b <= 0.04045 else ((b + 0.055) / 1.055) ** 2.4
    x = (r * 0.4124 + g * 0.3576 + b * 0.1805) / 0.95047
    y =  r * 0.2126 + g * 0.7152 + b * 0.0722
    z = (r * 0.0193 + g * 0.1192 + b * 0.9505) / 1.08883
    x, y, z = [v**(1/3) if v > 0.008856 else 7.787*v + 16/116 for v in [x, y, z]]
    result = (116*y - 16, 500*(x - y), 200*(y - z))
    _LAB_CACHE[key] = result
    return result

def delta_e(lab1, lab2):
    L1, a1, b1 = lab1; L2, a2, b2 = lab2
    C1 = math.sqrt(a1**2 + b1**2); C2 = math.sqrt(a2**2 + b2**2)
    C_avg = (C1 + C2) / 2.0; C_avg7 = C_avg**7
    G = 0.5 * (1.0 - math.sqrt(C_avg7 / (C_avg7 + 6103515625.0)))
    a1p = a1 * (1.0 + G);  a2p = a2 * (1.0 + G)
    C1p = math.sqrt(a1p**2 + b1**2); C2p = math.sqrt(a2p**2 + b2**2)
    def _h(ap, bp):
        if ap == 0.0 and bp == 0.0: return 0.0
        v = math.degrees(math.atan2(bp, ap))
        return v + 360.0 if v < 0.0 else v
    h1p = _h(a1p, b1); h2p = _h(a2p, b2)
    dLp = L2 - L1; dCp = C2p - C1p
    if C1p * C2p == 0.0: dhp = 0.0
    elif abs(h2p - h1p) <= 180.0: dhp = h2p - h1p
    elif h2p - h1p > 180.0: dhp = h2p - h1p - 360.0
    else: dhp = h2p - h1p + 360.0
    dHp = 2.0 * math.sqrt(C1p * C2p) * math.sin(math.radians(dhp / 2.0))
    Lp_avg = (L1 + L2) / 2.0; Cp_avg = (C1p + C2p) / 2.0
    if C1p * C2p == 0.0: hp_avg = h1p + h2p
    elif abs(h1p - h2p) <= 180.0: hp_avg = (h1p + h2p) / 2.0
    elif h1p + h2p < 360.0: hp_avg = (h1p + h2p + 360.0) / 2.0
    else: hp_avg = (h1p + h2p - 360.0) / 2.0
    T = (1.0 - 0.17 * math.cos(math.radians(hp_avg - 30.0))
         + 0.24 * math.cos(math.radians(2.0 * hp_avg))
         + 0.32 * math.cos(math.radians(3.0 * hp_avg + 6.0))
         - 0.20 * math.cos(math.radians(4.0 * hp_avg - 63.0)))
    SL = 1.0 + 0.015 * (Lp_avg - 50.0)**2 / math.sqrt(20.0 + (Lp_avg - 50.0)**2)
    SC = 1.0 + 0.045 * Cp_avg; SH = 1.0 + 0.015 * Cp_avg * T
    Cp_avg7 = Cp_avg**7
    RC = 2.0 * math.sqrt(Cp_avg7 / (Cp_avg7 + 6103515625.0))
    d_theta = 30.0 * math.exp(-((hp_avg - 275.0) / 25.0)**2)
    RT = -math.sin(math.radians(2.0 * d_theta)) * RC
    return math.sqrt((dLp/SL)**2 + (dCp/SC)**2 + (dHp/SH)**2 + RT*(dCp/SC)*(dHp/SH))

def seq_to_runs(sequence):
    if not sequence: return []
    runs = []
    cur, cnt = sequence[0], 1
    for x in sequence[1:]:
        if x == cur: cnt += 1
        else: runs.append((cur, cnt)); cur, cnt = x, 1
    runs.append((cur, cnt))
    return runs

def calc_cadence(sequence, layer_height):
    if not sequence or layer_height <= 0:
        return {}
    runs = seq_to_runs(sequence)
    counts = {}
    for fid, n in runs:
        counts[fid] = counts.get(fid, 0) + n
    total = sum(counts.values())
    if len(counts) == 1:
        fid = list(counts.keys())[0]
        return {fid: round(len(sequence) * layer_height, 3)}
    min_count = min(counts.values())
    result = {}
    for fid, cnt in counts.items():
        layers = round(cnt / min_count)
        layers = max(1, layers)
        result[fid] = round(layers * layer_height, 3)
    return result

def check_stripe_risk(sequence):
    if not sequence or len(sequence) < 2:
        return False, ""
    cycle = len(sequence)
    phase = (cycle // 2) + 1
    risky = (cycle % 2 == 0 and len(set(sequence)) == 2)
    if risky:
        return True, f"⚠ Stripe risk: cycle={cycle}, consider odd length"
    return False, f"✓ Pattern OK (cycle={cycle})"

MAX_SEQ_LEN = 48

def build_weighted_gradient_sequence(weights, max_len=48):
    n = len(weights)
    if n == 0: return []
    if n == 1: return [0] * min(max_len, 4)
    total = sum(weights)
    if total <= 0: return list(range(n))
    normed = [w / total for w in weights]
    counts = [max(1, round(w * max_len)) for w in normed]
    while sum(counts) > max_len:
        idx = max(range(n), key=lambda i: counts[i])
        counts[idx] -= 1
    while sum(counts) < max_len:
        idx = max(range(n), key=lambda i: normed[i] - counts[i]/max_len)
        counts[idx] += 1
    seq = []
    for i, c in enumerate(counts):
        seq.extend([i] * c)
    return seq


# ═══════════════════════════════════════════════════════════════════════════════
# TESTS
# ═══════════════════════════════════════════════════════════════════════════════

class TestHexRgb:
    def test_white(self):
        assert hex_to_rgb("#FFFFFF") == (255, 255, 255)

    def test_black(self):
        assert hex_to_rgb("#000000") == (0, 0, 0)

    def test_red(self):
        assert hex_to_rgb("#FF0000") == (255, 0, 0)

    def test_no_hash(self):
        assert hex_to_rgb("FF0000") == (255, 0, 0)

    def test_invalid_returns_grey(self):
        assert hex_to_rgb("ZZZZZZ") == (128, 128, 128)

    def test_short_returns_grey(self):
        assert hex_to_rgb("#FFF") == (128, 128, 128)

    def test_roundtrip(self):
        for hex_in in ["#1A2B3C", "#FF6600", "#00FF80"]:
            r, g, b = hex_to_rgb(hex_in)
            assert rgb_to_hex(r, g, b) == hex_in


class TestRgbToHex:
    def test_white(self):
        assert rgb_to_hex(255, 255, 255) == "#FFFFFF"

    def test_black(self):
        assert rgb_to_hex(0, 0, 0) == "#000000"

    def test_clamps_over(self):
        assert rgb_to_hex(300, 0, 0) == "#FF0000"

    def test_clamps_under(self):
        assert rgb_to_hex(-10, 128, 0) == "#008000"


class TestRgbToLab:
    def test_white_luminance(self):
        L, a, b = rgb_to_lab((255, 255, 255))
        assert abs(L - 100.0) < 1.0

    def test_black_luminance(self):
        L, a, b = rgb_to_lab((0, 0, 0))
        assert abs(L) < 1.0

    def test_neutral_grey_achromatic(self):
        L, a, b = rgb_to_lab((128, 128, 128))
        assert abs(a) < 2.0
        assert abs(b) < 2.0

    def test_cache_works(self):
        _LAB_CACHE.clear()
        result1 = rgb_to_lab((200, 100, 50))
        result2 = rgb_to_lab((200, 100, 50))
        assert result1 == result2
        assert (200, 100, 50) in _LAB_CACHE


class TestDeltaE:
    def test_identical_colors_zero(self):
        lab = rgb_to_lab((200, 100, 50))
        assert delta_e(lab, lab) == pytest.approx(0.0, abs=1e-6)

    def test_black_white_large(self):
        black = rgb_to_lab((0, 0, 0))
        white = rgb_to_lab((255, 255, 255))
        assert delta_e(black, white) > 50.0

    def test_similar_colors_small(self):
        c1 = rgb_to_lab((200, 100, 50))
        c2 = rgb_to_lab((202, 101, 51))
        assert delta_e(c1, c2) < 2.0

    def test_symmetric(self):
        lab1 = rgb_to_lab((255, 0, 0))
        lab2 = rgb_to_lab((0, 0, 255))
        assert abs(delta_e(lab1, lab2) - delta_e(lab2, lab1)) < 0.01

    def test_red_green_perceptibly_different(self):
        red = rgb_to_lab((220, 30, 30))
        green = rgb_to_lab((30, 180, 30))
        assert delta_e(red, green) > 50.0


class TestCalcCadence:
    def test_single_filament(self):
        seq = [1, 1, 1, 1]
        cad = calc_cadence(seq, 0.08)
        assert 1 in cad
        assert abs(cad[1] - round(4 * 0.08, 3)) < 0.001

    def test_equal_50_50(self):
        seq = [1, 2, 1, 2]
        cad = calc_cadence(seq, 0.08)
        assert cad[1] == pytest.approx(cad[2], abs=0.001)

    def test_2_to_1_ratio(self):
        # seq: T1 appears twice as often as T2 → T1 cadence = 2*lh, T2 = 1*lh
        seq = [1, 1, 2, 1, 1, 2]
        cad = calc_cadence(seq, 0.08)
        assert cad[1] == pytest.approx(cad[2] * 2, abs=0.01)

    def test_zero_layer_height(self):
        assert calc_cadence([1, 2], 0.0) == {}

    def test_empty_sequence(self):
        assert calc_cadence([], 0.08) == {}


class TestCheckStripeRisk:
    def test_even_cycle_two_filaments_risky(self):
        risk, msg = check_stripe_risk([1, 2, 1, 2])
        assert risk is True

    def test_odd_cycle_not_risky(self):
        risk, msg = check_stripe_risk([1, 2, 1, 2, 1])
        assert risk is False

    def test_three_filaments_not_risky(self):
        risk, msg = check_stripe_risk([1, 2, 3, 1, 2, 3])
        assert risk is False

    def test_single_filament_not_risky(self):
        risk, msg = check_stripe_risk([1, 1, 1, 1])
        assert risk is False

    def test_empty_not_risky(self):
        risk, msg = check_stripe_risk([])
        assert risk is False


class TestBuildWeightedGradient:
    def test_length_respects_max(self):
        seq = build_weighted_gradient_sequence([0.5, 0.5], max_len=10)
        assert len(seq) == 10

    def test_single_weight(self):
        seq = build_weighted_gradient_sequence([1.0], max_len=4)
        assert all(x == 0 for x in seq)

    def test_equal_weights_roughly_equal(self):
        seq = build_weighted_gradient_sequence([1, 1], max_len=10)
        assert seq.count(0) == 5
        assert seq.count(1) == 5

    def test_no_weights_returns_empty(self):
        seq = build_weighted_gradient_sequence([], max_len=10)
        assert seq == []

    def test_heavy_weight_more_layers(self):
        seq = build_weighted_gradient_sequence([3, 1], max_len=8)
        assert seq.count(0) > seq.count(1)

    def test_max_len_48(self):
        seq = build_weighted_gradient_sequence([1, 2, 3], max_len=48)
        assert len(seq) == 48
        assert all(x in (0, 1, 2) for x in seq)


class TestSeqToRuns:
    def test_basic(self):
        assert seq_to_runs([1, 1, 2, 2, 2, 1]) == [(1, 2), (2, 3), (1, 1)]

    def test_single(self):
        assert seq_to_runs([1]) == [(1, 1)]

    def test_empty(self):
        assert seq_to_runs([]) == []

    def test_alternating(self):
        assert seq_to_runs([1, 2, 1, 2]) == [(1, 1), (2, 1), (1, 1), (2, 1)]
