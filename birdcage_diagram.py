#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Birdcage Diagram (multi-slice Excel inputs) — Dash + Plotly

Key change vs single-table version:
- Each Excel file represents ONE slice.
- You provide multiple Excel files in the desired order: slice1, slice2, ...
- Category/level switching: choose the Category column in the UI (or via --category).
- Cross-level shift: element identity is the column immediately RIGHT of the selected Category column.
  (The --element-col is still used for fill-down exclusion and as the leaf-most column when Category is right before it.)


Expected per-slice Excel columns (example):
  Supergroup | Group | Hierarchy 3 | Hierarchy 4 | Element
- Leave blanks in hierarchy columns to use "fill-down" (向下继承). The script ffill()s
  all columns except Element.
- Element values should be unique within a slice (each element belongs to exactly one category).

Run (Windows / PyCharm parameters example):
  python birdcage_diagram.py --slice "slice1.xlsx" --slice "slice2.xlsx" --slice "slice3.xlsx" --slice "slice4.xlsx" --slice "slice5.xlsx" --slice "slice6.xlsx" --category "Hierarchy 4"

Dependencies:
  pip install pandas openpyxl numpy plotly dash kaleido
"""

from __future__ import annotations
import json
import re
import unicodedata
import base64
import uuid
import tempfile
from typing import List, Optional, Literal, Dict, Any
import math
import argparse
import dataclasses
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Set, Any
import plotly.graph_objects as go

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import dash
from dash import Dash, dcc, html, Input, Output, State
from dash.dependencies import ALL, MATCH
from dash.exceptions import PreventUpdate
import functools


def _estimate_text_width_px(text: str, font_size: float) -> float:
    """
    Very lightweight width estimator (no PIL).
    Tune k if your font looks wider/narrower.
    """
    k = 0.60
    n = len(text) if text is not None else 0
    if n <= 0:
        n = 1
    return float(n) * float(font_size) * k


def apply_group_label_clearance(
        group_order: list[str],
        group_x: dict[str, tuple[float, float]],
        block_x: dict[str, tuple[float, float]],
        group_to_blocks: dict[str, list[str]],
        *,
        left_hard_min_x: float,
        left_soft_min_x: float,
        label_text: dict[str, str],
        label_font_size: float,
        label_offset: float,
        gap: float,
) -> tuple[dict[str, tuple[float, float]], dict[str, tuple[float, float]]]:
    """
    Enforce horizontal room for group labels by shifting LEFT prefix blocks/frames first.
    If prefix cannot move further left (hits left_soft_min_x), push current+suffix RIGHT.

    Label box is:
        [gx0 - label_offset - text_width, gx0 - label_offset]
    Constraint:
        gx0 - label_offset - text_width >= prev_right + gap
    """
    group_x = dict(group_x)
    block_x = dict(block_x)

    def _shift_groups(gs: list[str], dx: float) -> None:
        if dx == 0.0:
            return
        for gg in gs:
            x0, x1 = group_x[gg]
            group_x[gg] = (x0 + dx, x1 + dx)
            for bb in group_to_blocks.get(gg, []):
                bx0, bx1 = block_x[bb]
                block_x[bb] = (bx0 + dx, bx1 + dx)

    def _min_x_of_groups(gs: list[str]) -> float:
        m = float("inf")
        for gg in gs:
            x0, _ = group_x[gg]
            if x0 < m:
                m = x0
        return m

    prev_right = left_soft_min_x

    for i, g in enumerate(group_order):
        gx0, _ = group_x[g]

        txt = label_text.get(g, g)
        w = _estimate_text_width_px(txt, label_font_size)

        label_right = gx0 - label_offset
        label_left = label_right - w

        required_left = prev_right + gap

        if label_left < required_left:
            need = required_left - label_left

            prefix = group_order[:i]
            if len(prefix) > 0:
                prefix_min_x = _min_x_of_groups(prefix)
                max_left_shift = prefix_min_x - left_soft_min_x
                if max_left_shift < 0.0:
                    max_left_shift = 0.0

                shift_left = need if need < max_left_shift else max_left_shift
                if shift_left > 0.0:
                    _shift_groups(prefix, -shift_left)
                    prev_right = prev_right - shift_left
                    need = need - shift_left

            if need > 0.0:
                suffix = group_order[i:]
                _shift_groups(suffix, need)
                gx0, _ = group_x[g]
                label_right = gx0 - label_offset
                label_left = label_right - w

        prev_right = group_x[g][1]

    all_min_x = float("inf")
    for x0, _ in group_x.values():
        if x0 < all_min_x:
            all_min_x = x0

    if all_min_x < left_hard_min_x:
        dx = left_hard_min_x - all_min_x
        _shift_groups(group_order, dx)

    return group_x, block_x


def _clamp01(x: float) -> float:
    if x < 0.0:
        return 0.0
    if x > 1.0:
        return 1.0
    return float(x)


def _is_hex_color(x: Any) -> bool:
    if not isinstance(x, str):
        return False
    if len(x) != 7 or (not x.startswith("#")):
        return False
    for ch in x[1:]:
        if ch not in "0123456789abcdefABCDEF":
            return False
    return True


def _hex_to_rgb(hex_color: str) -> Tuple[int, int, int]:
    s = str(hex_color).strip()
    if s.startswith("#"):
        s = s[1:]
    if len(s) != 6:
        raise ValueError(f"Expected #RRGGBB, got: {hex_color}")
    r = int(s[0:2], 16)
    g = int(s[2:4], 16)
    b = int(s[4:6], 16)
    return r, g, b


def _rgb_to_hex(r: int, g: int, b: int) -> str:
    r = int(max(0, min(255, r)))
    g = int(max(0, min(255, g)))
    b = int(max(0, min(255, b)))
    return f"#{r:02X}{g:02X}{b:02X}"


def lighten_hex(hex_color: str, amount: float) -> str:
    """
    Mix with white by `amount` in [0,1]. Keeps hue, reduces saturation.
    """
    a = _clamp01(amount)
    if not _is_hex_color(hex_color):
        return str(hex_color)
    r, g, b = _hex_to_rgb(hex_color)
    r2 = int(round(r + (255 - r) * a))
    g2 = int(round(g + (255 - g) * a))
    b2 = int(round(b + (255 - b) * a))
    return _rgb_to_hex(r2, g2, b2)


def darken_hex(hex_color: str, amount: float) -> str:
    """
    Mix with black by `amount` in [0,1]. Keeps hue, increases perceived strength.
    """
    a = _clamp01(amount)
    if not _is_hex_color(hex_color):
        return str(hex_color)
    r, g, b = _hex_to_rgb(hex_color)
    r2 = int(round(r * (1.0 - a)))
    g2 = int(round(g * (1.0 - a)))
    b2 = int(round(b * (1.0 - a)))
    return _rgb_to_hex(r2, g2, b2)


def _is_rgba_color(x: Any) -> bool:
    """Check if x is a valid rgba(...) color string."""
    if not isinstance(x, str):
        return False
    s = x.strip().lower()
    return s.startswith("rgba(") and s.endswith(")")


def _normalize_hex_color(s: Any, fallback: str = "#999999") -> str:
    """Normalize a hex color string. Accepts '#RRGGBB' or 'RRGGBB'.

    Returns *fallback* if invalid.

    Note: This is intentionally lightweight and local, to avoid adding new dependencies.
    """
    if not isinstance(s, str):
        return fallback
    t = s.strip()
    if not t:
        return fallback
    if not t.startswith('#'):
        t = '#' + t
    if len(t) != 7:
        return fallback
    hexd = set('0123456789abcdefABCDEF')
    if any(ch not in hexd for ch in t[1:]):
        return fallback
    return t


def _parse_rgba(rgba_str: str) -> Tuple[int, int, int, float]:
    """Parse rgba(r,g,b,a) string and return (r, g, b, a)."""
    s = rgba_str.strip()
    if s.lower().startswith("rgba(") and s.endswith(")"):
        inner = s[5:-1]
        parts = [p.strip() for p in inner.split(",")]
        if len(parts) == 4:
            r = int(float(parts[0]))
            g = int(float(parts[1]))
            b = int(float(parts[2]))
            a = float(parts[3])
            return r, g, b, a
    raise ValueError(f"Invalid rgba string: {rgba_str}")


def _to_rgba_str(r: int, g: int, b: int, a: float) -> str:
    """Convert r, g, b, a to rgba(...) string."""
    r = int(max(0, min(255, r)))
    g = int(max(0, min(255, g)))
    b = int(max(0, min(255, b)))
    a = max(0.0, min(1.0, a))
    return f"rgba({r},{g},{b},{a})"


def darken_rgba(rgba_color: str, amount: float) -> str:
    """
    Darken an rgba color by mixing with black.
    """
    a = _clamp01(amount)
    if not _is_rgba_color(rgba_color):
        return str(rgba_color)
    try:
        r, g, b, alpha = _parse_rgba(rgba_color)
        r2 = int(round(r * (1.0 - a)))
        g2 = int(round(g * (1.0 - a)))
        b2 = int(round(b * (1.0 - a)))
        return _to_rgba_str(r2, g2, b2, alpha)
    except Exception:
        return str(rgba_color)


def darken_color(color: str, amount: float) -> str:
    """
    Darken a color (supports both hex and rgba formats).
    """
    if _is_hex_color(color):
        return darken_hex(color, amount)
    elif _is_rgba_color(color):
        return darken_rgba(color, amount)
    else:
        return str(color)


def _lighten_color_map(m: Dict[str, str], amount: float) -> Dict[str, str]:
    return {k: lighten_hex(v, amount) for k, v in m.items()}


# Global palette tuning knobs (requested behavior)
_BASE_LIGHTEN = 0.0  # colors below are already the desired pastel/fill defaults
_SELECTED_DARKEN = 0.45  # selected core becomes stronger (darker)

# Default fill colors for block types (used as-is since _BASE_LIGHTEN=0).
_BLOCK_COLORS_STRONG: Dict[str, str] = {
    "ReOut": "#FFE699",
    "ReIn": "#B3C8E5",
    "EnBirth": "#C5E0B5",
    "ExBirth": "#C5E0B5",
    "HyBirth": "#C5E0B5",
    "Birth": "#C5E0B5",
    "Death": "#FF7D7D",
    "Exo": "#FFFFFF",
    "Persis": "#E8E8E8",
    "Unknown": "#CCCCCC",
}

_BAND_COLORS_STRONG: Dict[str, str] = {
    "Inflow": "#D5E8FB",
    "Outflow": "#FEE7CB",
    "Merge": "#EE9A4F",
    "Split": "#8846EC",
    "SpMe": "#00897B",
    "Inheri": "#D7DDDF",
    "Unknown": "#9E9E9E",
}

# Ordered list of band types shown in the UI panel (excluding Unknown)
BAND_TYPE_UI_ORDER: List[str] = ["Outflow", "Inflow", "Split", "Merge", "SpMe", "Inheri"]
BLOCK_TYPE_UI_ORDER: List[str] = ["All", "ReOut", "ReIn", "Birth", "EnBirth", "ExBirth", "HyBirth", "Death", "Exo",
                                  "Persis"]


# 1) Style and layout parameters


@dataclass
class StyleConfig:
    # Geometry (data units)
    block_width: float = 1.5
    block_height: float = 5.0
    inner_gap: float = 0.0
    outer_gap: float = 0.8
    layer_gap: float = 5.0

    # Exo positioning (kept as-is)
    exo_gap: float = 2.0

    # --- Font sizes ---
    # Sized to match the visual density the user calibrated against
    # (Firefox at 30% browser zoom, where the figure read as "right" —
    # all text small enough to fit fully inside blocks instead of being
    # truncated to 5-character ellipses, while shapes and proportions
    # remain readable). At a 100% browser zoom (what reviewers will use)
    # these absolute px values reproduce that same visual.
    # Earlier values were ~3× larger (block 20, ui 20, tooltip 18, etc.);
    # those produced the over-sized look the user reported as "too big in
    # Chrome/Edge". The factor used here is ~0.35 — derived from direct
    # screenshot comparison of the same figure at the two browser zooms.
    block_text_size: int = 7   # Block labels
    ui_font_size: int = 7      # UI labels, dropdowns, etc.
    tooltip_font_size: int = 18 # Hover tooltip — reference px @ 1920 viewport (rendered via vw, scales with browser width)
    button_font_size: int = 7  # Buttons

    # Exo-toggle "button" inside plot (per-slice)
    exo_toggle_marker_size: int = 11  # clickable square size (was 30)
    exo_toggle_text_size: int = 7     # text on the square
    exo_toggle_gap: float = 0.35  # distance from exo block left edge to toggle button center

    # Exo hidden marker (small "E" square) — make it light when visible
    exo_toggle_marker_fill: str = "#E8E8E8"  # light gray fill for restore button
    exo_toggle_marker_line: str = "#BBBBBB"  # border color
    exo_toggle_text_color: str = "#999999"  # text color

    # Viewport padding for symmetric centering (REQUEST #2)
    viewport_pad_x: float = 0.8
    viewport_pad_y: float = 0.8

    # Band geometry
    band_curve_strength: float = 0.35
    band_samples: int = 20
    band_min_opacity: float = 0.08
    band_default_opacity: float = 1.0
    band_highlight_opacity: float = 0.85

    # NEW: fixed straight-band width (data units)
    band_width_ratio: float = 0.06  # thickness = block_width * 0.06 = 0.09
    band_stroke_width: float = 0.18  # outline width (px) — scaled to match font scale (was 0.5)

    # Block visual
    block_border_width: float = 0.35   # was 1.0; scaled with font defaults
    block_text_color_default: str = "#111111"

    # Group / enclosure
    enclosure_padding_x: float = 0.06
    enclosure_padding_y: float = 0.06
    enclosure_line_width: float = 0.35   # was 1.0; scaled
    enclosure_opacity: float = 0.35
    enclosure_pad_step_x: float = 0.04
    enclosure_pad_step_y: float = 0.04

    # Enclosure label (shown near top border midpoint)
    enclosure_label_text_size: int = 7   # was 20; scaled
    enclosure_label_text_color: str = "#111111"

    # Label placement (DATA units): anchor on border y1; text is drawn ABOVE via textposition="top center"
    enclosure_label_offset_y_base: float = 0.00
    enclosure_label_offset_y_step: float = 0.00

    # --- Group label placed on the LEFT side of group enclosure ---
    enclosure_group_label_offset_x: float = 0.7  # label anchor = rx0 - offset_x
    enclosure_label_clearance_x: float = 0.04  # extra whitespace when expanding OUTER enclosures to the left

    # Side-mask tuning (cover the LEFT border segment behind group label)
    enclosure_label_mask_side_in_x: float = 0.03  # mask extends inside the border by this amount
    enclosure_label_mask_side_pad_y: float = 0.02  # vertical padding around the side label mask

    # White label box (mask) sizing helpers (DATA units)
    # NOTE: actual box height is computed from font size; these are only minimums / small margins.
    enclosure_label_mask_up_y: float = 0.02
    enclosure_label_mask_down_y: float = 0.01

    # Mask width estimation (DATA units): approx text width = n_chars * (block_width * ratio) + padding
    enclosure_label_mask_char_w_ratio: float = 0.07
    enclosure_label_mask_pad_x: float = 0.02

    # When labels are ON, reserve vertical space by lifting OUTER enclosures upward (top edge only)
    enclosure_label_clearance_y: float = 0.03
    enclosure_label_height_scale: float = 0.55
    enclosure_label_height_min_ratio: float = 0.55

    # Annotation background (white label box)
    enclosure_label_bgcolor: str = "#FFFFFF"
    enclosure_label_bordercolor: str = "#FFFFFF"
    enclosure_label_borderpad: int = 1

    # 更保守的 group 左侧标签宽度安全系数（避免中文/混合字符估算偏小）
    enclosure_group_label_width_fudge: float = 1.0
    # 侧边标签估算字符宽比例（原来 0.07 容易偏小，建议提高）
    enclosure_label_mask_char_w_ratio: float = 0.095

    # Gap between sibling enclosures at the SAME level (horizontal)
    enclosure_sibling_gap_x: float = 0.22  # base gap
    enclosure_sibling_gap_step_x: float = 0.10  # extra gap for higher (outer) levels

    # --- Group-level enclosure (closest-to-category; glued blocks) ---
    group_enclosure_line_width: float = 1.75  # was 5.0; scaled to match font defaults
    group_enclosure_opacity: float = 1.0
    # Make group enclosure slightly OUTSIDE blocks so it stays visible even when drawn below traces,
    # and so the label background does not overlap blocks.
    group_enclosure_pad_x: float = 0.03
    group_enclosure_pad_y: float = 0.03

    # --- Freeze Exo x so it never shifts when main-body width changes (NEW) ---
    exo_anchor_x: Optional[float] = None

    # Strength thresholds
    tau_major: float = 0.60
    tau_intermediate: float = 0.25

    # Highlight dim
    dim_opacity: float = 0.08
    # Highlight intensity
    core_block_opacity: float = 1.0
    related_block_opacity: float = 0.62
    core_band_opacity: float = 0.95
    related_band_opacity: float = 0.75

    core_block_border_width: float = 3.0
    related_block_border_width: float = 1.6
    core_band_line_width: float = 2.6
    related_band_line_width: float = 1.2

    # Background click-catcher (almost invisible but clickable)
    bg_click_alpha: float = 0.001

    # Axes
    show_axes: bool = False

    # Block colors by label (PASTEL BASE)
    block_colors: Dict[str, str] = dataclasses.field(
        default_factory=lambda: _lighten_color_map(_BLOCK_COLORS_STRONG, _BASE_LIGHTEN))

    # Block text colors (keep readable on fills)
    block_text_colors: Dict[str, str] = dataclasses.field(
        default_factory=lambda: {k: ("#FFFFFF" if k == "Death" else "#111111") for k in _BLOCK_COLORS_STRONG.keys()})

    # Band colors by label (PASTEL BASE)
    band_colors: Dict[str, str] = dataclasses.field(
        default_factory=lambda: _lighten_color_map(_BAND_COLORS_STRONG, _BASE_LIGHTEN))

    # ========== Block Style Per Type (NEW) ==========
    # block_widths: per-type block width override (data units)
    block_widths_by_type: Dict[str, float] = dataclasses.field(default_factory=dict)
    # block_heights: per-type block height override (data units)
    block_heights_by_type: Dict[str, float] = dataclasses.field(default_factory=dict)
    # block_fill_colors: per-type fill color override (#RRGGBB)
    block_fill_colors_by_type: Dict[str, str] = dataclasses.field(default_factory=dict)
    # block_fill_opacities: per-type fill opacity override (0.0 - 1.0)
    block_fill_opacities_by_type: Dict[str, float] = dataclasses.field(default_factory=dict)
    # block_border_colors: per-type border color override (#RRGGBB)
    block_border_colors_by_type: Dict[str, str] = dataclasses.field(default_factory=dict)
    # block_border_opacities: per-type border opacity override (0.0 - 1.0)
    block_border_opacities_by_type: Dict[str, float] = dataclasses.field(default_factory=dict)
    # block_border_radii: per-type corner radius override (list of 4: [top-left, top-right, bottom-right, bottom-left])
    block_border_radii_by_type: Dict[str, List[float]] = dataclasses.field(default_factory=dict)
    # block_line_styles: per-type line style override ("solid", "dash", "dot", "dashdot")
    block_line_styles_by_type: Dict[str, str] = dataclasses.field(default_factory=dict)
    # block_text_fonts: per-type text font override
    block_text_fonts_by_type: Dict[str, str] = dataclasses.field(default_factory=dict)
    # block_text_sizes: per-type text size override
    block_text_sizes_by_type: Dict[str, int] = dataclasses.field(default_factory=dict)
    # block_text_colors_by_type: per-type text color override (#RRGGBB)
    block_text_colors_by_type: Dict[str, str] = dataclasses.field(default_factory=dict)
    # block_text_aligns: per-type text alignment ("left", "center", "right", "justify")
    block_text_aligns_by_type: Dict[str, str] = dataclasses.field(default_factory=dict)
    # block_text_rotations_by_type: per-type text rotation angle in degrees
    block_text_rotations_by_type: Dict[str, int] = dataclasses.field(default_factory=dict)
    # block_line_spacings_by_type: per-type line spacing in absolute px offset (0 = default)
    block_line_spacings_by_type: Dict[str, float] = dataclasses.field(default_factory=dict)
    # block_border_widths_by_type: per-type border width override
    block_border_widths_by_type: Dict[str, float] = dataclasses.field(default_factory=dict)

    # ========== Block Text Padding (distance from text to each edge, data units) ==========
    block_text_pad_l: float = 0.0
    block_text_pad_r: float = 0.0
    block_text_pad_t: float = 0.0
    block_text_pad_b: float = 0.0

    # Fractional safety inset applied to the block's text area on top of the
    # per-side pad_* values above. DEFAULT 0.0 — i.e. disabled by default,
    # matching the v44 baseline that the user confirmed produced the desired
    # visual density. The mid-word hyphenation in _wrap_text (added in v48)
    # already prevents char-direction overflow, and the _CHAR_W=0.63 estimate
    # already accounts for line-direction safety, so this extra inset just
    # wastes block capacity. Kept as a config knob so users who want a more
    # padded look (e.g. for export / print) can opt in via cfg, but the
    # default behaviour is the v44 one (no inset, no wasted space).
    block_text_safety_inset_frac: float = 0.0

    # ========== Group Enclosure Style (NEW) ==========
    # group_style: dict with keys: line_width, opacity, color, radii (list of 4), line_style
    group_style: Dict[str, Any] = dataclasses.field(default_factory=lambda: {
        "line_width": 1.75,   # was 5.0; scaled to match font defaults
        "opacity": 1.0,
        "color": "#000000",
        "radii": [0, 0, 0, 0],
        "line_style": "solid",
        "label_size": 7,      # was 20; scaled
    })

    # ========== Supergroup Enclosure Styles (NEW) ==========
    # supergroup_styles: dict keyed by supergroup level (1, 2, 3, ...)
    # each value is a dict with keys:
    #   pad_left, pad_right, pad_top, pad_bottom (distance from group enclosure)
    #   fill_color, fill_opacity, border_color, border_opacity, radii (list of 4), line_style
    supergroup_styles: Dict[int, Dict[str, Any]] = dataclasses.field(default_factory=dict)


EXO_NAME = "__EXO__"


# 2) Reading slice files, fill-down, parsing

@functools.lru_cache(maxsize=32)
def read_slice_excel(path: str, sheet_name=0) -> pd.DataFrame:
    """Read one slice table from a file path.

    Supported:
      - .xlsx / .xls / .ods  -> pandas.read_excel (single sheet or named sheet)
      - .csv                 -> pandas.read_csv (UTF-8 with fallback)
    """
    p = str(path)
    ext = str(Path(p).suffix).lower()

    if ext in [".xlsx", ".xls", ".ods"]:
        df = pd.read_excel(p, sheet_name=sheet_name)
    elif ext == ".csv":
        try:
            df = pd.read_csv(p, encoding="utf-8")
        except Exception:
            df = pd.read_csv(p, encoding="latin1")
    else:
        raise ValueError(f"Unsupported slice file extension: {ext}. Supported: .xlsx, .xls, .ods, .csv")

    if df.shape[1] < 2:
        raise ValueError(f"Slice file has too few columns: {path}")
    return df


def read_all_sheets(path: str) -> tuple:
    """Read all sheets from a multi-sheet Excel file.
    Returns (list of DataFrames, list of sheet names).
    """
    p = str(path)
    ext = str(Path(p).suffix).lower()
    if ext not in [".xlsx", ".xls", ".ods"]:
        raise ValueError(f"Multi-sheet mode only supports .xlsx/.xls/.ods, got: {ext}")
    xl = pd.ExcelFile(p)
    sheet_names = xl.sheet_names
    dfs = [xl.parse(s) for s in sheet_names]
    return dfs, sheet_names


def fill_down_per_slice(df: pd.DataFrame, exclude_cols: List[str]) -> pd.DataFrame:
    """
    Fill-down (ffill) all columns except those in exclude_cols.

    Cross-level shift rule:
    - Always exclude the leaf-most Element ID column given by --element-col.
    - Do not exclude the right-adjacent "element identity" column in general, because when the Category
      is shifted to a higher level, the right-adjacent column is typically a hierarchy column that uses
      fill-down blanks in the raw Excel. Excluding it will leave many blanks and those rows will be dropped.
    """

    out = df.copy()

    # normalize exclude list, keep order stable
    ex: List[str] = []
    for c in exclude_cols:
        if c not in ex:
            ex.append(c)

    missing = [c for c in ex if c not in out.columns]
    if missing:
        raise ValueError(f"Fill-down exclusion columns missing: {missing}. Columns: {list(out.columns)}")

    non_fill = set(ex)
    fill_cols = [c for c in out.columns if c not in non_fill]
    if fill_cols:
        out[fill_cols] = out[fill_cols].ffill()

    return out


def validate_columns_for_choice(cols: List[str], category_col: str, element_col: str) -> List[str]:
    """
    Returns agg_cols (nearest-first) = reversed(cols before category), excluding element_col (rightmost col).
    element_col is always cols[-1] (the rightmost column).
    """
    if category_col not in cols:
        raise ValueError(f"Category column '{category_col}' not found. Available: {cols}")
    # element_col is always the rightmost column - no need to validate by name
    element_col = cols[-1]
    cat_idx = cols.index(category_col)
    left_cols = [c for c in cols[:cat_idx] if c != element_col]
    agg_cols = list(reversed(left_cols))  # nearest-first
    return agg_cols


def right_adjacent_column(cols: List[str], category_col: str) -> str:
    """
    Return the column immediately to the RIGHT of category_col.
    This is the "element identity" column under cross-level shift.
    """
    if category_col not in cols:
        raise ValueError(f"Category column '{category_col}' not found. Available: {cols}")
    idx = cols.index(category_col)
    if idx >= len(cols) - 1:
        raise ValueError(
            f"Category column '{category_col}' has no right-adjacent column. "
            f"Please choose a category column that is NOT the last column."
        )
    return cols[idx + 1]


# 3) Internal model: slices, blocks, groups

@dataclass(frozen=True)
class BlockKey:
    slice_id: str
    cat_name: str  # EXO_NAME for Exo


@dataclass(frozen=True)
class BandKey:
    src_slice: str
    src_cat: str
    dst_slice: str
    dst_cat: str


@dataclass
class TreeNode:
    name: str
    level: str  # "root", "group", "supergroup_k", "leaf"
    start: int
    end: int
    children: List["TreeNode"]
    leaf_cat: Optional[str] = None

    def is_leaf(self) -> bool:
        return self.leaf_cat is not None


@dataclass
class SliceModel:
    slice_id: str
    categories_in_order: List[str]  # excludes Exo
    cat_to_elems: Dict[str, Set[str]]
    level_membership: Dict[str, Dict[str, str]]  # agg_col -> (cat -> val)
    interval_tree: TreeNode
    agg_cols: List[str]


def _build_intervals(categories: List[str], cat_to_val: Dict[str, str]) -> List[Tuple[int, int, str]]:
    intervals: List[Tuple[int, int, str]] = []
    if not categories:
        return intervals
    cur_val = cat_to_val.get(categories[0], "")
    start = 0
    for idx, cat in enumerate(categories):
        v = cat_to_val.get(cat, "")
        if idx == 0:
            cur_val = v
            continue
        if v != cur_val:
            intervals.append((start, idx - 1, str(cur_val)))
            start = idx
            cur_val = v
    intervals.append((start, len(categories) - 1, str(cur_val)))
    return intervals


def _bbox_union(bbs: List[Tuple[float, float, float, float]]) -> Tuple[float, float, float, float]:
    x0 = min(bb[0] for bb in bbs)
    x1 = max(bb[1] for bb in bbs)
    y0 = min(bb[2] for bb in bbs)
    y1 = max(bb[3] for bb in bbs)
    return (x0, x1, y0, y1)


def _bbox_expand(bb: Tuple[float, float, float, float], pad_x: float, pad_y: float) -> Tuple[
    float, float, float, float]:
    x0, x1, y0, y1 = bb
    return (x0 - pad_x, x1 + pad_x, y0 - pad_y, y1 + pad_y)


def rounded_rect_polygon_coords(
        x0: float, y0: float, x1: float, y1: float,
        radii: List[float],
        segments: int = 8
) -> Tuple[List[float], List[float]]:
    """
    Generate polygon coordinates for a rounded rectangle.

    Args:
        x0, y0: bottom-left corner
        x1, y1: top-right corner
        radii: [top-left, top-right, bottom-right, bottom-left] corner radii
        segments: number of segments per corner arc

    Returns:
        (xs, ys): lists of x and y coordinates forming a closed polygon
    """
    import math

    # Ensure radii is a list of 4 values, handle None values
    if not radii or len(radii) != 4:
        radii = [0, 0, 0, 0]

    # Convert to floats, treating None as 0
    def safe_float(r):
        if r is None:
            return 0.0
        try:
            return max(0.0, float(r))
        except (ValueError, TypeError):
            return 0.0

    r_tl, r_tr, r_br, r_bl = [safe_float(r) for r in radii]

    # If all radii are 0, return simple rectangle
    if r_tl == 0 and r_tr == 0 and r_br == 0 and r_bl == 0:
        return [x0, x1, x1, x0, x0], [y0, y0, y1, y1, y0]

    width = x1 - x0
    height = y1 - y0

    # Scale radii if they're too large for the rectangle
    # Each corner radius should not exceed half the width or height
    max_r_x = width / 2.0
    max_r_y = height / 2.0

    # Scale down if necessary
    scale = 1.0
    if r_tl + r_tr > width:
        scale = min(scale, width / (r_tl + r_tr))
    if r_bl + r_br > width:
        scale = min(scale, width / (r_bl + r_br))
    if r_tl + r_bl > height:
        scale = min(scale, height / (r_tl + r_bl))
    if r_tr + r_br > height:
        scale = min(scale, height / (r_tr + r_br))

    if scale < 1.0:
        r_tl *= scale
        r_tr *= scale
        r_br *= scale
        r_bl *= scale

    # Clamp individual radii
    r_tl = min(r_tl, max_r_x, max_r_y)
    r_tr = min(r_tr, max_r_x, max_r_y)
    r_br = min(r_br, max_r_x, max_r_y)
    r_bl = min(r_bl, max_r_x, max_r_y)

    xs = []
    ys = []

    # Start from bottom-left, go clockwise
    # Bottom edge (left to right)
    xs.append(x0 + r_bl)
    ys.append(y0)
    xs.append(x1 - r_br)
    ys.append(y0)

    # Bottom-right corner
    if r_br > 0:
        cx, cy = x1 - r_br, y0 + r_br
        for i in range(segments + 1):
            angle = -math.pi / 2 + (math.pi / 2) * i / segments
            xs.append(cx + r_br * math.cos(angle))
            ys.append(cy + r_br * math.sin(angle))
    else:
        xs.append(x1)
        ys.append(y0)

    # Right edge (bottom to top)
    xs.append(x1)
    ys.append(y1 - r_tr)

    # Top-right corner
    if r_tr > 0:
        cx, cy = x1 - r_tr, y1 - r_tr
        for i in range(segments + 1):
            angle = 0 + (math.pi / 2) * i / segments
            xs.append(cx + r_tr * math.cos(angle))
            ys.append(cy + r_tr * math.sin(angle))
    else:
        xs.append(x1)
        ys.append(y1)

    # Top edge (right to left)
    xs.append(x0 + r_tl)
    ys.append(y1)

    # Top-left corner
    if r_tl > 0:
        cx, cy = x0 + r_tl, y1 - r_tl
        for i in range(segments + 1):
            angle = math.pi / 2 + (math.pi / 2) * i / segments
            xs.append(cx + r_tl * math.cos(angle))
            ys.append(cy + r_tl * math.sin(angle))
    else:
        xs.append(x0)
        ys.append(y1)

    # Left edge (top to bottom)
    xs.append(x0)
    ys.append(y0 + r_bl)

    # Bottom-left corner
    if r_bl > 0:
        cx, cy = x0 + r_bl, y0 + r_bl
        for i in range(segments + 1):
            angle = math.pi + (math.pi / 2) * i / segments
            xs.append(cx + r_bl * math.cos(angle))
            ys.append(cy + r_bl * math.sin(angle))
    else:
        xs.append(x0)
        ys.append(y0)

    # Close the polygon
    xs.append(xs[0])
    ys.append(ys[0])

    return xs, ys


def rounded_rect_svg_path(
        x0: float, y0: float, x1: float, y1: float,
        radii: List[float]
) -> str:
    """
    Generate a Plotly-compatible SVG path string for a rounded rectangle.

    Notes:
    - Plotly layout.shapes path rendering does not reliably support SVG elliptical-arc
      commands (A). When arcs are not supported, corners appear as chamfers.
    - To guarantee rounded-looking corners for group/supergroup enclosures, we
      approximate each quarter-circle with a polyline (sampled points) and emit
      only M/L/Z commands.
    """
    xs, ys = rounded_rect_polygon_coords(x0, y0, x1, y1, radii, segments=12)
    if not xs or not ys or len(xs) != len(ys):
        return f"M {x0},{y0} L {x1},{y0} L {x1},{y1} L {x0},{y1} Z"

    # Ensure closure
    if xs[0] != xs[-1] or ys[0] != ys[-1]:
        xs = list(xs) + [xs[0]]
        ys = list(ys) + [ys[0]]

    cmd = [f"M {xs[0]},{ys[0]}"]
    for x, y in zip(xs[1:], ys[1:]):
        cmd.append(f"L {x},{y}")
    cmd.append("Z")
    return " ".join(cmd)


def build_enclosure_shapes_for_slice(
        sl: SliceModel,
        cats: List[str],
        layout: LayoutInfo,
        cfg: StyleConfig,
        ppu_x: float = 30.0,   # effective pixels-per-data-unit (at initial display zoom)
) -> Tuple[List[dict], List[dict], List[str], List[int], List[int], List[int]]:
    """
    Bottom-up enclosure bboxes with symmetric margins.
    - Innermost drawn level: bbox = union(block bboxes) + base padding.
    - Outer levels: bbox = union(child enclosure bboxes) + step padding (incremental).
    - At each level, cap pad_x symmetrically so sibling borders keep at least
      cfg.enclosure_sibling_gap_x (+ depth * step) horizontal whitespace.
    Notes:
    - We NEVER shrink below inner bbox (pad_x >= 0), so containment is preserved.
    """
    sid = sl.slice_id

    # agg_cols is nearest-first (closest-to-category first).
    # We now ALSO draw the closest-to-category level (the glued-group level)
    # with a bold enclosure, even when the group has only one block.
    outer_cols = list(reversed(sl.agg_cols))  # outermost -> ... -> closest-to-category
    if not outer_cols:
        return [], [], [], [], [], [], None

    # Draw ALL levels including the closest-to-category (group) level
    levels_draw = outer_cols  # outermost -> ... -> closest-to-category

    # for bottom-up computation: innermost-drawn -> ... -> outermost-drawn
    levels_inner_to_outer = list(reversed(levels_draw))

    # Closest-to-category col = glued-group level
    group_col = sl.agg_cols[0]

    # intervals per level (in terms of category indices)
    intervals_by_col: Dict[str, List[Tuple[int, int, str]]] = {}
    for col in levels_draw:
        cat_to_val = sl.level_membership.get(col, {})
        intervals_by_col[col] = _build_intervals(cats, cat_to_val)

    def union_blocks(s_idx: int, e_idx: int) -> Tuple[float, float, float, float]:
        bbs = [layout.block_bbox[BlockKey(sid, cats[k])] for k in range(s_idx, e_idx + 1)]
        return _bbox_union(bbs)

    def cap_symmetric_pad_x(
            inner_bbs: List[Tuple[float, float, float, float]],
            desired_pad_x: float,
            sibling_gap_x: float,
    ) -> List[float]:
        """
        Compute per-box pad_x (applied to both left/right),
        capped by neighbor whitespace so adjacent borders keep >= sibling_gap_x.
        """
        n = len(inner_bbs)
        if n == 0:
            return []

        g = float(max(0.0, sibling_gap_x))
        pads = [float(max(0.0, desired_pad_x))] * n

        for i in range(n):
            left_gap = float("inf") if i == 0 else (inner_bbs[i][0] - inner_bbs[i - 1][1])
            right_gap = float("inf") if i == n - 1 else (inner_bbs[i + 1][0] - inner_bbs[i][1])

            cap_left = float("inf") if left_gap == float("inf") else max(0.0, (left_gap - g) / 2.0)
            cap_right = float("inf") if right_gap == float("inf") else max(0.0, (right_gap - g) / 2.0)

            pads[i] = min(pads[i], cap_left, cap_right)

        return pads

    # computed bboxes for each drawn level (same order as intervals_by_col[col])
    bboxes_by_col: Dict[str, List[Tuple[float, float, float, float]]] = {}

    # Get group and supergroup styles from config (needed for UI padding in bbox computation)
    group_style = getattr(cfg, "group_style", {}) or {}
    supergroup_styles = getattr(cfg, "supergroup_styles", {}) or {}

    # 1) innermost drawn level: from blocks
    col0 = levels_inner_to_outer[0]
    inner0: List[Tuple[float, float, float, float]] = []
    for (s_idx, e_idx, _val) in intervals_by_col[col0]:
        inner0.append(union_blocks(s_idx, e_idx))

    # Group level should be very tight but NOT coincident with block border:
    # add a tiny pad so border stays visible and label box does not overlap blocks.
    if col0 == group_col:
        bboxes_by_col[col0] = [_bbox_expand(bb, float(cfg.group_enclosure_pad_x), float(cfg.group_enclosure_pad_y)) for
                               bb in inner0]
    else:
        gap0 = cfg.enclosure_sibling_gap_x
        pads0 = cap_symmetric_pad_x(inner0, cfg.enclosure_padding_x, gap0)
        bbs0 = [_bbox_expand(bb, px, cfg.enclosure_padding_y) for bb, px in zip(inner0, pads0)]
        bboxes_by_col[col0] = bbs0

    # 2) outer levels: union child enclosures, then expand by STEP padding (incremental)
    for depth in range(1, len(levels_inner_to_outer)):
        col = levels_inner_to_outer[depth]
        child_col = levels_inner_to_outer[depth - 1]

        cur_intervals = intervals_by_col[col]
        child_intervals = intervals_by_col[child_col]
        child_bbs = bboxes_by_col[child_col]

        inner: List[Tuple[float, float, float, float]] = []
        for (s_idx, e_idx, _val) in cur_intervals:
            contained_child = []
            for ci, (cs, ce, _cval) in enumerate(child_intervals):
                if cs >= s_idx and ce <= e_idx:
                    contained_child.append(child_bbs[ci])

            inner.append(_bbox_union(contained_child) if contained_child else union_blocks(s_idx, e_idx))

        # sibling gap can grow for outer levels
        gap = cfg.enclosure_sibling_gap_x + depth * cfg.enclosure_sibling_gap_step_x

        pads = cap_symmetric_pad_x(inner, cfg.enclosure_pad_step_x, gap)
        cur_bbs = [_bbox_expand(bb, px, cfg.enclosure_pad_step_y) for bb, px in zip(inner, pads)]

        # Apply UI-controlled padding INTO bboxes_by_col so parent levels
        # see the ACTUAL drawn borders when computing their union.
        # This ensures the gap between a parent and child border is preserved
        # when the child's padding changes.
        level_idx = sl.agg_cols.index(col) if col in sl.agg_cols else -1
        if col != group_col and level_idx > 0:
            sg_style = supergroup_styles.get(level_idx, {})

            def _sf(v, d=0.0):
                try:
                    return float(v) if v is not None and v != "" else d
                except Exception:
                    return d

            ui_pl = _sf(sg_style.get("pad_left"))
            ui_pr = _sf(sg_style.get("pad_right"))
            ui_pt = _sf(sg_style.get("pad_top"))
            ui_pb = _sf(sg_style.get("pad_bottom"))
            cur_bbs = [(x0 - ui_pl, x1 + ui_pr, y0 - ui_pb, y1 + ui_pt)
                       for (x0, x1, y0, y1) in cur_bbs]

        bboxes_by_col[col] = cur_bbs

    # 3) emit shapes with controlled intra-shape z-order.
    # We keep ALL borders first, then ALL label background masks.
    # This guarantees that (for any level) the white label background can sit on top of
    # inner borders (e.g., supergroup background above group border).
    border_shapes: List[dict] = []
    mask_shapes: List[dict] = []
    label_specs: List[dict] = []  # each: {"x":..., "y":..., "text":...}
    label_level_keys: List[str] = []  # per label spec
    mask_shape_indices: List[int] = []  # per label spec (white label box shape index)
    border_shape_indices: List[int] = []  # per label spec (enclosure border shape index)
    label_depths: List[int] = []  # per label spec (0=group, 1=supergroup1, ...)

    def _get_dash_pattern(line_style: str) -> Optional[str]:
        if line_style == "dash":
            return "dash"
        elif line_style == "dot":
            return "dot"
        elif line_style == "dashdot":
            return "dashdot"
        return None  # solid

    def _hex_to_rgba(hex_color: str, opacity: float) -> str:
        hex_color = str(hex_color).strip()
        if not hex_color.startswith("#"):
            hex_color = "#" + hex_color
        if len(hex_color) != 7:
            return f"rgba(0,0,0,{opacity})"
        try:
            r = int(hex_color[1:3], 16)
            g = int(hex_color[3:5], 16)
            b = int(hex_color[5:7], 16)
            return f"rgba({r},{g},{b},{opacity})"
        except Exception:
            return f"rgba(0,0,0,{opacity})"

    def _safe_float(v: Any, default: float) -> float:
        try:
            if v is None or v == "" or v == "None":
                return float(default)
            return float(v)
        except Exception:
            return float(default)

    for level_rank, col in enumerate(levels_draw):  # outermost -> inner
        intervals = intervals_by_col[col]
        bbs = bboxes_by_col[col]

        level_idx = sl.agg_cols.index(col) if col in sl.agg_cols else -1
        level_key = "group" if level_idx == 0 else f"supergroup{level_idx}"

        if col == group_col:
            # Apply group style
            lw = _safe_float(group_style.get("line_width"), cfg.group_enclosure_line_width)
            op = _safe_float(group_style.get("opacity"), cfg.group_enclosure_opacity)
            border_color = str(group_style.get("color") or "#000000")
            line_style = str(group_style.get("line_style") or "solid")
            radii_raw = group_style.get("radii")
            radii = list(radii_raw) if radii_raw and isinstance(radii_raw, (list, tuple)) and len(radii_raw) == 4 else [
                0, 0, 0, 0]
            # Ensure radii has no None values
            radii = [0 if r is None else r for r in radii]
            fill_color = "rgba(0,0,0,0)"  # Group has no fill
            fill_opacity = 0.0
        else:
            # Apply supergroup style
            sg_level = level_idx  # 1, 2, 3, ...
            sg_style = supergroup_styles.get(sg_level, {})

            default_lw = max(0.5, cfg.enclosure_line_width * (0.9 ** level_rank))
            lw = _safe_float(sg_style.get("border_width"), default_lw)
            default_op = cfg.enclosure_opacity * (0.85 ** level_rank)
            op = _safe_float(sg_style.get("border_opacity"), default_op)
            border_color = str(sg_style.get("border_color") or "#000000")
            line_style = str(sg_style.get("line_style") or "solid")
            radii_raw = sg_style.get("radii")
            radii = list(radii_raw) if radii_raw and isinstance(radii_raw, (list, tuple)) and len(radii_raw) == 4 else [
                0, 0, 0, 0]
            # Ensure radii has no None values
            radii = [0 if r is None else r for r in radii]
            fill_color = str(sg_style.get("fill_color") or "#FFFFFF")
            fill_opacity = _safe_float(sg_style.get("fill_opacity"), 0.0)

        for (_s_idx, _e_idx, _val), bb in zip(intervals, bbs):
            x0, x1, y0, y1 = bb
            # UI padding is already included in bboxes_by_col (applied during bbox computation)
            # so we do NOT add it again here.

            # Layering requirement:
            # - Group enclosure border must be ABOVE blocks and bands.
            # - Supergroup enclosures remain BELOW bands/blocks.
            layer_here = "above" if (col == group_col) else "below"

            # 1) enclosure border with style
            bidx = len(border_shapes)

            # Create line dict with optional dash pattern
            line_dict = dict(color=_hex_to_rgba(border_color, op), width=lw)
            dash_pattern = _get_dash_pattern(line_style)
            if dash_pattern:
                line_dict["dash"] = dash_pattern

            # Create fill color with opacity
            if fill_opacity > 0.001:
                fill_rgba = _hex_to_rgba(fill_color, fill_opacity)
            else:
                fill_rgba = "rgba(0,0,0,0)"

            border_shapes.append(dict(
                type="rect",
                x0=x0, x1=x1, y0=y0, y1=y1,
                line=line_dict,
                fillcolor=fill_rgba,
                opacity=1.0,  # opacity is baked into colors
                layer=layer_here,
            ))

            # Check if we need rounded corners and convert to path
            if any((r if r is not None else 0) > 0 for r in radii):
                # Scale radii from pixel-like values to data units
                block_w = float(cfg.block_width)
                block_h = float(cfg.block_height)
                scale_factor = min(block_w, block_h) / 50.0  # 50 is max radius in UI
                radii_scaled = [(r if r is not None else 0) * scale_factor for r in radii]

                # Generate SVG path for rounded rectangle
                svg_path = rounded_rect_svg_path(x0, y0, x1, y1, radii_scaled)

                # Replace the last added shape with path version
                border_shapes[-1] = dict(
                    type="path",
                    path=svg_path,
                    x0=x0, x1=x1, y0=y0, y1=y1,
                    xref="x", yref="y",
                    line=line_dict,
                    fillcolor=fill_rgba,
                    opacity=1.0,
                    layer=layer_here,
                )

            # 2) "line-break mask" on the top border (initially hidden, toggled with label)
            xmid = (x0 + x1) / 2.0
            txt_raw = str(_val)

            # Tight mask width + ellipsis rule:
            # - Background should be only slightly wider than the rendered text.
            # - Text width is capped at 50% of the enclosure's top edge length; overflow uses an ellipsis.
            # char_w: width of one average character in data units.
            # Use font-size / ppu_x × 0.72 (conservative; covers wide chars A,W,M ≈ 0.70-0.85 of font height).
            top_len = float(x1) - float(x0)
            _font_px = float(cfg.enclosure_label_text_size)
            char_w = _font_px / max(float(ppu_x), 1.0) * 0.72
            pad_x = float(cfg.enclosure_label_mask_pad_x) * 0.35

            max_w = float(top_len) * 0.50   # 50% of enclosure width — safe margin

            def _est_text_w_data(s: str) -> float:
                n = max(1, len(s))
                return float(n) * float(char_w) + 2.0 * float(pad_x)

            txt = txt_raw
            # Pre-truncate the displayed text to keep the line-break mask
            # tight within the enclosure top edge. SKIP for group labels:
            #   - Group labels are rotated 90°, so their available space is
            #     the enclosure HEIGHT, not its top-edge width — truncating
            #     against width here vastly under-counts available chars
            #     for tall, narrow group enclosures.
            #   - The mask itself is hidden for group labels in
            #     apply_enclosure_label_visibility (it sets shapes[midx].visible
            #     = False on the group branch), so the mask-sizing rationale
            #     for this truncation does not apply.
            #   - Pre-truncating here would clip the input that _wrap_text_mod
            #     later sees in apply_enclosure_label_visibility, defeating
            #     its (correct, height-based, hyphenation-aware) truncation.
            # For supergroup labels (top-center, non-rotated), keep this
            # truncation as-is — the mask IS shown and needs to fit the top
            # edge.
            if str(level_key) != "group":
                if max_w > 0.0 and _est_text_w_data(txt_raw) > max_w:
                    denom = max(float(char_w), 1e-9)
                    max_chars = int(max(1, math.floor((max_w - 2.0 * float(pad_x)) / denom)))
                    if max_chars <= 1:
                        txt = "…"
                    else:
                        keep = max(1, int(max_chars) - 1)  # reserve 1 char for ellipsis
                        txt = str(txt_raw)[:keep].rstrip() + "…"

            n_chars = max(1, len(txt))
            half_w = 0.5 * float(n_chars) * float(char_w) + float(pad_x)

            mx0 = xmid - half_w
            mx1 = xmid + half_w
            mx0 = max(float(x0), float(mx0))
            mx1 = min(float(x1), float(mx1))
            if mx1 <= mx0:
                mx0 = xmid - 1e-6
                mx1 = xmid + 1e-6

            my0 = float(y1) - float(cfg.enclosure_label_mask_down_y)
            my1 = float(y1) + float(cfg.enclosure_label_mask_up_y)

            # Prevent label background masks from intruding into the block area.
            # We keep the mask mostly ABOVE the enclosure top border, with only a small downward extent.
            mask_depth = 0 if str(level_key) == "group" else int(str(level_key).replace("supergroup", ""))
            down_factor = 0.10 if int(mask_depth) == 0 else 0.05
            my0 = float(y1) - float(cfg.enclosure_label_mask_down_y) * float(down_factor)

            midx_local = len(mask_shapes)
            # White label background (initially hidden, toggled with label)
            # NOTE: must be ABOVE group enclosure border, but BELOW bands.
            # We achieve this by:
            #   - placing masks AFTER all borders in `layout.shapes`
            #   - letting bands be drawn above via trace/shapes order (handled in build_figure)
            mask_shapes.append(dict(
                type="rect",
                x0=mx0, x1=mx1, y0=my0, y1=my1,
                line=dict(color="rgba(0,0,0,0)", width=0),
                fillcolor="#FFFFFF",
                opacity=1.0,
                layer="below",
                visible=False,
            ))
            mask_shape_indices.append(int(midx_local))
            border_shape_indices.append(int(bidx))

            # 3) label spec (drawn as a TRACE so it can be placed under bands by trace order)
            level_depth = 0 if level_key == "group" else int(level_key.replace("supergroup", ""))
            spec_entry: dict = {"x": float(xmid), "y": float(y1), "text": str(txt),
                                "orig_text": str(txt_raw)}
            if txt != txt_raw:
                # Truncated: store full name for JS tooltip
                spec_entry["full_text"] = str(txt_raw)
            label_specs.append(spec_entry)
            label_level_keys.append(str(level_key))
            label_depths.append(int(level_depth))

    shapes = list(border_shapes) + list(mask_shapes)
    mask_offset = len(border_shapes)
    mask_shape_indices = [int(mask_offset + int(i)) for i in mask_shape_indices]

    # Compute outermost enclosure union bbox for this slice.
    # UI padding is already included in bboxes_by_col, so no additional adjustment needed.
    outermost_col = levels_inner_to_outer[-1] if levels_inner_to_outer else None
    outermost_bbox: Optional[Tuple[float, float, float, float]] = None
    if outermost_col and outermost_col in bboxes_by_col and bboxes_by_col[outermost_col]:
        outermost_bbox = _bbox_union(bboxes_by_col[outermost_col])

    return shapes, label_specs, label_level_keys, mask_shape_indices, border_shape_indices, label_depths, outermost_bbox


def build_interval_tree(categories: List[str], memberships: Dict[str, Dict[str, str]], agg_cols: List[str]) -> TreeNode:
    leaf_nodes = [
        TreeNode(name=cat, level="leaf", start=i, end=i, children=[], leaf_cat=cat)
        for i, cat in enumerate(categories)
    ]

    if not agg_cols:
        return TreeNode(name="root", level="root", start=0, end=max(0, len(categories) - 1), children=leaf_nodes)

    current_nodes: List[TreeNode] = leaf_nodes

    # IMPORTANT: agg_cols is nearest-first (closest-to-category first).
    # Build the tree from inner to outer so that:
    # leaf -> group -> supergroup_1 -> supergroup_2 -> ...
    for lvl_idx, col in enumerate(agg_cols):
        cat_to_val = memberships.get(col, {})
        intervals = _build_intervals(categories, cat_to_val)

        new_nodes: List[TreeNode] = []
        for (s, e, val) in intervals:
            # Children at the next-inner level must be fully contained in [s, e]
            child_slice = [n for n in current_nodes if (n.start >= s and n.end <= e)]

            if not child_slice:
                raise ValueError(
                    f"Interval tree construction failed at level '{col}' for interval [{s}, {e}]='{val}'. "
                    f"This usually indicates that the hierarchy levels are built in the wrong order, "
                    f"or the membership values are inconsistent with the current category order."
                )

            node = TreeNode(
                name=str(val),
                level=("group" if lvl_idx == 0 else f"supergroup_{lvl_idx}"),
                start=s,
                end=e,
                children=child_slice,
                leaf_cat=None,
            )
            new_nodes.append(node)

        current_nodes = new_nodes

    return TreeNode(
        name="root",
        level="root",
        start=0,
        end=max(0, len(categories) - 1),
        children=current_nodes,
    )


def filter_nested_agg_cols(
        agg_cols: List[str],
        cat_meta: Dict[str, Dict[str, str]],
        *,
        slice_id: str = "",
        on_violation: Literal["error", "drop_inner"] = "error",
) -> List[str]:
    """
    Keep only a nested (functional) chain in agg_cols (nearest-first).

    Nested requirement for consecutive levels:
      agg_cols[k] (inner) must map to exactly ONE value of agg_cols[k+1] (outer)
      across all categories.

    If violated:
      - on_violation="error": raise ValueError (recommended if you want GUARANTEED continuity)
      - on_violation="drop_inner": drop the inner level (previous behavior)
    """
    cols = list(agg_cols)
    cats = list(cat_meta.keys())

    k = 0
    while k < len(cols) - 1:
        inner = cols[k]
        outer = cols[k + 1]

        mp: Dict[str, Set[str]] = defaultdict(set)
        for cat in cats:
            iv = "" if (cat_meta.get(cat, {}).get(inner, "") is None) else str(cat_meta[cat].get(inner, ""))
            ov = "" if (cat_meta.get(cat, {}).get(outer, "") is None) else str(cat_meta[cat].get(outer, ""))
            mp[iv].add(ov)

        conflicts: Dict[str, List[str]] = {}
        for iv, ovs in mp.items():
            if len(ovs) > 1:
                conflicts[str(iv)] = sorted([str(x) for x in ovs])

        if conflicts:
            if on_violation == "error":
                sid = f" '{slice_id}'" if str(slice_id).strip() != "" else ""
                # show a small sample to keep the message readable
                items = list(conflicts.items())
                items.sort(key=lambda t: (t[0], ",".join(t[1])))
                sample = items[:12]
                sample_txt = "\n".join([f"  inner='{iv}' -> outers={ovs}" for iv, ovs in sample])
                raise ValueError(
                    f"Non-nested hierarchy detected in slice{sid} between inner column '{inner}' and outer column '{outer}'.\n"
                    f"This makes it impossible to keep BOTH levels continuous simultaneously.\n"
                    f"This is a DATA constraint violation (hierarchy is not a strict tree), not a layout/reorder bug.\n"
                    f"Fix your Excel so each '{inner}' value belongs to exactly one '{outer}' value.\n"
                    f"Conflicts (sample):\n{sample_txt}"
                )

            # on_violation == "drop_inner": keep older behavior
            cols.pop(k)
            continue

        k += 1

    return cols


def build_slice_model(df: pd.DataFrame, slice_id: str, category_col: str, element_col: str,
                      agg_cols: List[str]) -> SliceModel:
    cat_to_elems: Dict[str, Set[str]] = defaultdict(set)

    # Drop rows without element-id (right-adjacent column)
    df2 = df.copy()
    df2 = df2[df2[element_col].notna()].copy()
    df2[element_col] = df2[element_col].astype(str).str.strip()
    df2 = df2[df2[element_col] != ""].copy()

    # Collect per-category meta (stable ordering + validate consistent memberships)
    cat_first_pos: Dict[str, int] = {}
    cat_meta: Dict[str, Dict[str, str]] = {}

    for pos, (_, r) in enumerate(df2.iterrows()):
        cat_val = r.get(category_col, None)
        elem_val = r.get(element_col, None)
        if pd.isna(cat_val) or pd.isna(elem_val):
            continue

        cat = str(cat_val).strip()
        elem = str(elem_val).strip()
        if elem == "":
            continue

        if cat not in cat_first_pos:
            cat_first_pos[cat] = int(pos)
            cat_meta[cat] = {}

        cat_to_elems[cat].add(elem)

        for c in agg_cols:
            v = r.get(c, "")
            vv = "" if pd.isna(v) else str(v).strip()
            if c not in cat_meta[cat]:
                cat_meta[cat][c] = vv
            else:
                if cat_meta[cat][c] != vv:
                    raise ValueError(
                        f"Inconsistent membership in slice '{slice_id}': "
                        f"category '{cat}' has multiple values in column '{c}': "
                        f"'{cat_meta[cat][c]}' vs '{vv}'."
                    )

    # NEW: filter agg_cols to keep only a nested chain (nearest-first)
    agg_cols2 = filter_nested_agg_cols(list(agg_cols), cat_meta, slice_id=str(slice_id), on_violation="error")

    # Finalize level_membership from cat_meta (ONLY for agg_cols2)
    level_membership: Dict[str, Dict[str, str]] = {c: {} for c in agg_cols2}
    for c in agg_cols2:
        for cat, meta in cat_meta.items():
            level_membership[c][cat] = meta.get(c, "")

    # Order categories by hierarchy keys (outer -> inner), then by first appearance (stable)
    if len(cat_meta) == 0:
        cat_order: List[str] = []
    else:
        outer_to_inner = list(reversed(agg_cols2))  # outermost first

        def _sort_key(cat: str):
            return tuple(cat_meta[cat].get(col, "") for col in outer_to_inner) + (cat_first_pos.get(cat, 0),)

        cat_order = sorted(list(cat_meta.keys()), key=_sort_key)

    # Validate element uniqueness across categories (allow duplicates within same category because sets dedup)
    all_elems: List[str] = []
    for _c, es in cat_to_elems.items():
        all_elems.extend(list(es))
    if len(all_elems) != len(set(all_elems)):
        raise ValueError(
            f"Duplicate element IDs detected within slice '{slice_id}'. "
            f"Each element must belong to exactly one category per slice."
        )

    tree = build_interval_tree(cat_order, level_membership, agg_cols2)
    return SliceModel(
        slice_id=str(slice_id),
        categories_in_order=cat_order,
        cat_to_elems=dict(cat_to_elems),
        level_membership=level_membership,
        interval_tree=tree,
        agg_cols=list(agg_cols2),
    )


# 4) Alignment matrices and event labeling

@dataclass
class StepModel:
    src_slice: str
    dst_slice: str
    src_cats: List[str]  # includes EXO at 0
    dst_cats: List[str]  # includes EXO at 0
    A: np.ndarray  # counts


def compute_step_alignment(src: SliceModel, dst: SliceModel) -> StepModel:
    src_cats = [EXO_NAME] + src.categories_in_order
    dst_cats = [EXO_NAME] + dst.categories_in_order
    src_idx = {c: i for i, c in enumerate(src_cats)}
    dst_idx = {c: j for j, c in enumerate(dst_cats)}

    src_elem_to_cat = {e: c for c, es in src.cat_to_elems.items() for e in es}
    dst_elem_to_cat = {e: c for c, es in dst.cat_to_elems.items() for e in es}

    elems_union = set(src_elem_to_cat.keys()) | set(dst_elem_to_cat.keys())
    A = np.zeros((len(src_cats), len(dst_cats)), dtype=int)
    for e in elems_union:
        s_cat = src_elem_to_cat.get(e, EXO_NAME)
        d_cat = dst_elem_to_cat.get(e, EXO_NAME)
        if s_cat == EXO_NAME and d_cat == EXO_NAME:
            continue
        A[src_idx[s_cat], dst_idx[d_cat]] += 1

    return StepModel(src_slice=src.slice_id, dst_slice=dst.slice_id, src_cats=src_cats, dst_cats=dst_cats, A=A)


def compute_all_steps(slices: List[SliceModel]) -> List[StepModel]:
    return [compute_step_alignment(a, b) for a, b in zip(slices[:-1], slices[1:])]


def _primary_block_label(label_set: Set[str]) -> str:
    if "ReOut" in label_set:
        return "ReOut"
    if "ReIn" in label_set:
        return "ReIn"
    for b in ["EnBirth", "ExBirth", "HyBirth", "Birth"]:
        if b in label_set:
            return b
    if "Death" in label_set:
        return "Death"
    if "Exo" in label_set:
        return "Exo"
    if "Persis" in label_set:
        return "Persis"
    return "Unknown"


def _primary_band_label(label_set: Set[str]) -> str:
    if "Outflow" in label_set:
        return "Outflow"
    if "Inflow" in label_set:
        return "Inflow"
    if "SpMe" in label_set:
        return "SpMe"
    if "Split" in label_set:
        return "Split"
    if "Merge" in label_set:
        return "Merge"
    if "Inheri" in label_set:
        return "Inheri"
    return "Unknown"


# --- NEW: separate priority lists for BLOCK vs BAND event types ---
_BLOCK_EVENT_PRIORITY = [
    "ReOut", "ReIn",
    "EnBirth", "ExBirth", "HyBirth", "Birth",
    "Death",
    "Persis",
    "Exo",
    "Unknown",
]

_BAND_EVENT_PRIORITY = [
    "Outflow",
    "Inflow",
    "SpMe",
    "Split",
    "Merge",
    "Inheri",
    "Unknown",
]

_BLOCK_EVENT_SET = set(_BLOCK_EVENT_PRIORITY)
_BAND_EVENT_SET = set(_BAND_EVENT_PRIORITY)


def format_block_event_text(labels: Set[str]) -> str:
    display = set(labels)

    # If a Birth subtype is known, do not show the generic "Birth"
    if ("EnBirth" in display) or ("ExBirth" in display) or ("HyBirth" in display):
        if "Birth" in display:
            display.remove("Birth")

    ordered = [x for x in _BLOCK_EVENT_PRIORITY if x in display]
    return "; ".join(ordered) if ordered else "Unknown"


def format_band_event_text(labels: Set[str]) -> str:
    ordered = [x for x in _BAND_EVENT_PRIORITY if x in labels]
    return "; ".join(ordered) if ordered else "Unknown"


def label_blocks_and_bands(
        slices: List[SliceModel],
        steps: List[StepModel],
        cfg: StyleConfig
) -> Tuple[Dict[BlockKey, Set[str]], Dict[BandKey, Set[str]]]:
    block_labels: Dict[BlockKey, Set[str]] = defaultdict(set)
    band_labels: Dict[BandKey, Set[str]] = defaultdict(set)

    # Exo blocks
    for sl in slices:
        block_labels[BlockKey(sl.slice_id, EXO_NAME)].update({"Exo", "Persis"})

    name_set = {sl.slice_id: set(sl.categories_in_order) for sl in slices}
    elems_set = {sl.slice_id: {c: set(es) for c, es in sl.cat_to_elems.items()} for sl in slices}

    for step in steps:
        src_id, dst_id = step.src_slice, step.dst_slice
        src_names = name_set[src_id]
        dst_names = name_set[dst_id]
        A = step.A
        De = {i: list(np.where(A[i, :] > 0)[0]) for i in range(A.shape[0])}
        So = {j: list(np.where(A[:, j] > 0)[0]) for j in range(A.shape[1])}

        src_cats = step.src_cats
        dst_cats = step.dst_cats

        # Death/Persis for sources
        for i, c in enumerate(src_cats):
            if c == EXO_NAME:
                continue
            key = BlockKey(src_id, c)
            if c in dst_names:
                block_labels[key].add("Persis")
            else:
                block_labels[key].add("Death")

        # Birth subtype for destinations
        for j, c in enumerate(dst_cats):
            if c == EXO_NAME:
                continue
            key = BlockKey(dst_id, c)
            if c not in src_names:
                block_labels[key].add("Birth")
                src_indices = [h for h in So[j] if int(A[h, j]) > 0]
                has_exo = (0 in src_indices)
                has_insys = any((h != 0) for h in src_indices)
                if has_insys and not has_exo:
                    block_labels[key].add("EnBirth")
                elif (not has_insys) and has_exo:
                    block_labels[key].add("ExBirth")
                elif has_insys and has_exo:
                    block_labels[key].add("HyBirth")
            else:
                block_labels[key].add("Persis")

        # Rename (ReOut/ReIn), strict one-to-one with full element sets
        # Conditions (BL1/BL2): λ_{t,i} ∉ Λ_{t+1} AND c_{t,i}=c_{t+1,j} AND λ_{t+1,j} ∉ Λ_t
        for i, src_cat in enumerate(src_cats):
            if src_cat == EXO_NAME or src_cat in dst_names:  # condition 1: λ_{t,i} ∉ Λ_{t+1}
                continue
            dests = [j for j in De[i] if j != 0]
            if len(dests) != 1:
                continue
            j = dests[0]
            dst_cat = dst_cats[j]
            if dst_cat in src_names:  # condition 3: λ_{t+1,j} ∉ Λ_t
                continue
            srcs = [h for h in So[j] if h != 0]
            if len(srcs) != 1 or srcs[0] != i:
                continue
            src_elems = elems_set[src_id].get(src_cat, set())
            dst_elems = elems_set[dst_id].get(dst_cat, set())
            shared = int(A[i, j])
            if shared == len(src_elems) == len(dst_elems):  # condition 2: identical element sets
                block_labels[BlockKey(src_id, src_cat)].add("ReOut")
                block_labels[BlockKey(dst_id, dst_cat)].add("ReIn")

        # Band labels
        # Phase 1: Inflow / Outflow
        for i, src_cat in enumerate(src_cats):
            for j, dst_cat in enumerate(dst_cats):
                v = int(A[i, j])
                if v <= 0:
                    continue
                if src_cat == EXO_NAME and dst_cat == EXO_NAME:
                    continue
                bk = BandKey(src_id, src_cat, dst_id, dst_cat)
                if src_cat == EXO_NAME and dst_cat != EXO_NAME:
                    band_labels[bk].add("Inflow")
                elif src_cat != EXO_NAME and dst_cat == EXO_NAME:
                    band_labels[bk].add("Outflow")

        # Phase 2: Split — source (non-Exo) with |Det(i) \ {Exo}| > 1
        # Per user spec: Exo destinations do NOT count toward the split threshold.
        for i, src_cat in enumerate(src_cats):
            if src_cat == EXO_NAME:
                continue
            src_dsts = [k for k in De[i] if int(A[i, k]) > 0]
            # Count only non-Exo destinations for the split condition
            non_exo_dsts = [k for k in src_dsts if dst_cats[k] != EXO_NAME]
            if len(non_exo_dsts) > 1:
                for j in src_dsts:
                    dst_cat = dst_cats[j]
                    if src_cat == EXO_NAME and dst_cat == EXO_NAME:
                        continue
                    bk = BandKey(src_id, src_cat, dst_id, dst_cat)
                    band_labels[bk].add("Split")

        # Phase 3: Merge — dest (non-Exo) with |So(j) \ {Exo}| > 1
        # Per user spec: Exo sources do NOT count toward the merge threshold.
        for j, dst_cat in enumerate(dst_cats):
            if dst_cat == EXO_NAME:
                continue
            dst_srcs = [h for h in So[j] if int(A[h, j]) > 0]
            # Count only non-Exo sources for the merge condition
            non_exo_srcs = [h for h in dst_srcs if src_cats[h] != EXO_NAME]
            if len(non_exo_srcs) > 1:
                for h in dst_srcs:
                    src_cat2 = src_cats[h]
                    if src_cat2 == EXO_NAME and dst_cat == EXO_NAME:
                        continue
                    bk = BandKey(src_id, src_cat2, dst_id, dst_cat)
                    band_labels[bk].add("Merge")

        # Phase 4: SpMe — both Split and Merge
        for i, src_cat in enumerate(src_cats):
            for j, dst_cat in enumerate(dst_cats):
                v = int(A[i, j])
                if v <= 0:
                    continue
                if src_cat == EXO_NAME and dst_cat == EXO_NAME:
                    continue
                bk = BandKey(src_id, src_cat, dst_id, dst_cat)
                if "Split" in band_labels[bk] and "Merge" in band_labels[bk]:
                    band_labels[bk].add("SpMe")

        # Phase 5: Inheri — fallback for non-Exo bands with no Split/Merge
        for i, src_cat in enumerate(src_cats):
            for j, dst_cat in enumerate(dst_cats):
                v = int(A[i, j])
                if v <= 0:
                    continue
                if src_cat == EXO_NAME or dst_cat == EXO_NAME:
                    continue
                bk = BandKey(src_id, src_cat, dst_id, dst_cat)
                if not (band_labels[bk] & {"Split", "Merge", "SpMe"}):
                    band_labels[bk].add("Inheri")

        # Strength buckets
        for j, dst_cat in enumerate(dst_cats):
            if j == 0:
                continue
            incoming = [(i, int(A[i, j])) for i in range(A.shape[0]) if int(A[i, j]) > 0]
            if len(incoming) <= 1:
                continue
            total = sum(v for _, v in incoming)
            if total <= 0:
                continue
            for i, v in incoming:
                src_cat = src_cats[i]
                bk = BandKey(src_id, src_cat, dst_id, dst_cat)
                contrib = v / total
                if contrib >= cfg.tau_major:
                    band_labels[bk].add("Major")
                elif contrib >= cfg.tau_intermediate:
                    band_labels[bk].add("Intermediate")
                else:
                    band_labels[bk].add("Minor")

        for i, src_cat in enumerate(src_cats):
            if i == 0:
                continue
            outgoing = [(j, int(A[i, j])) for j in range(A.shape[1]) if int(A[i, j]) > 0]
            if len(outgoing) <= 1:
                continue
            total = sum(v for _, v in outgoing)
            if total <= 0:
                continue
            for j, v in outgoing:
                dst_cat = dst_cats[j]
                bk = BandKey(src_id, src_cat, dst_id, dst_cat)
                contrib = v / total
                if contrib >= cfg.tau_major:
                    band_labels[bk].add("Major")
                elif contrib >= cfg.tau_intermediate:
                    band_labels[bk].add("Intermediate")
                else:
                    band_labels[bk].add("Minor")

    # (BL6) All categories in the last slice θ_T are assigned Persis
    if slices:
        last_sl = slices[-1]
        for c in last_sl.categories_in_order:
            block_labels[BlockKey(last_sl.slice_id, c)].add("Persis")

    return dict(block_labels), dict(band_labels)


# 5) Within-layer ordering (interval-tree reorder + sweep)

def _desc_weight(node: TreeNode, w_leaf: Dict[str, float]) -> float:
    if node.is_leaf():
        return float(w_leaf.get(node.leaf_cat or "", 1.0))
    return float(sum(_desc_weight(ch, w_leaf) for ch in node.children))


def interval_tree_reorder(root: TreeNode, categories: List[str], p_leaf: Dict[str, float], w_leaf: Dict[str, float]) -> \
List[str]:
    scores: Dict[int, float] = {}

    def post(n: TreeNode) -> float:
        if n.is_leaf():
            s = float(p_leaf.get(n.leaf_cat or "", 0.0))
            scores[id(n)] = s
            return s
        child_scores = [post(ch) for ch in n.children]
        weights = [_desc_weight(ch, w_leaf) for ch in n.children]
        denom = sum(weights)
        if denom <= 0:
            denom = float(len(child_scores)) if child_scores else 1.0
            weights = [1.0] * len(child_scores)
        s = sum(w * cs for w, cs in zip(weights, child_scores)) / denom
        scores[id(n)] = float(s)
        return float(s)

    def sort_td(n: TreeNode) -> None:
        if n.is_leaf():
            return
        n.children.sort(key=lambda ch: (scores.get(id(ch), 0.0), int(getattr(ch, "start", 0))))
        for ch in n.children:
            sort_td(ch)

    def expand(n: TreeNode, out: List[str]) -> None:
        if n.is_leaf():
            out.append(n.leaf_cat or "")
            return
        for ch in n.children:
            expand(ch, out)

    post(root)
    sort_td(root)
    out: List[str] = []
    expand(root, out)

    keep = set(categories)
    uniq: List[str] = []
    seen: Set[str] = set()
    for c in out:
        if c in keep and c not in seen:
            uniq.append(c)
            seen.add(c)
    for c in categories:
        if c not in seen:
            uniq.append(c)
    return uniq


def sweep_reorder(
        slices: List[SliceModel],
        steps: List[StepModel],
        K_max: int = 10,
        m: int = 2,
        delta: float = 0.01
) -> Dict[str, List[str]]:
    order = {sl.slice_id: list(sl.categories_in_order) for sl in slices}
    w = {sl.slice_id: {c: float(max(1, len(sl.cat_to_elems.get(c, set())))) for c in sl.categories_in_order} for sl in
         slices}
    step_by_pair = {(st.src_slice, st.dst_slice): st for st in steps}

    consec = 0
    for _k in range(K_max):
        prev = {sid: list(o) for sid, o in order.items()}

        # forward
        for idx in range(1, len(slices)):
            cur = slices[idx]
            prv = slices[idx - 1]
            st = step_by_pair[(prv.slice_id, cur.slice_id)]
            prev_pos = {cat: p for p, cat in enumerate(order[prv.slice_id])}
            pos0 = {cat: p for p, cat in enumerate(cur.categories_in_order)}

            p_leaf: Dict[str, float] = {}
            for cat in cur.categories_in_order:
                if cat not in st.dst_cats:
                    p_leaf[cat] = float(pos0.get(cat, 0))
                    continue
                j = st.dst_cats.index(cat)
                denom = 0
                num = 0.0
                for i, src_cat in enumerate(st.src_cats):
                    if i == 0:
                        continue
                    v = int(st.A[i, j])
                    if v <= 0:
                        continue
                    denom += v
                    num += v * float(prev_pos.get(src_cat, 0.0))
                p_leaf[cat] = float(num / denom) if denom > 0 else float(pos0.get(cat, 0))

            order[cur.slice_id] = interval_tree_reorder(cur.interval_tree, order[cur.slice_id], p_leaf, w[cur.slice_id])

        # backward
        for idx in range(len(slices) - 2, -1, -1):
            cur = slices[idx]
            nxt = slices[idx + 1]
            st = step_by_pair[(cur.slice_id, nxt.slice_id)]
            nxt_pos = {cat: p for p, cat in enumerate(order[nxt.slice_id])}
            pos0 = {cat: p for p, cat in enumerate(cur.categories_in_order)}

            p_leaf: Dict[str, float] = {}
            for cat in cur.categories_in_order:
                if cat not in st.src_cats:
                    p_leaf[cat] = float(pos0.get(cat, 0))
                    continue
                i = st.src_cats.index(cat)
                denom = 0
                num = 0.0
                for j, dst_cat in enumerate(st.dst_cats):
                    if j == 0:
                        continue
                    v = int(st.A[i, j])
                    if v <= 0:
                        continue
                    denom += v
                    num += v * float(nxt_pos.get(dst_cat, 0.0))
                p_leaf[cat] = float(num / denom) if denom > 0 else float(pos0.get(cat, 0))

            order[cur.slice_id] = interval_tree_reorder(cur.interval_tree, order[cur.slice_id], p_leaf, w[cur.slice_id])

        disp_sum = 0.0
        disp_n = 0
        for sl in slices:
            sid = sl.slice_id
            pos_old = {c: i for i, c in enumerate(prev[sid])}
            pos_new = {c: i for i, c in enumerate(order[sid])}
            for c in set(prev[sid]) & set(order[sid]):
                disp_sum += abs(pos_new[c] - pos_old[c])
                disp_n += 1
        disp = disp_sum / (disp_n if disp_n else 1)
        if disp < delta:
            consec += 1
        else:
            consec = 0
        if consec >= m:
            break

    return order


# 6) Multi-layer collapse (x,y anchors)

def collapse_steps(
        slices: List[SliceModel],
        steps: List[StepModel],
        x_idx: int,
        y_idx: int
) -> Tuple[List[SliceModel], List[StepModel], Optional[StepModel]]:
    if x_idx < 0 or y_idx >= len(slices) or x_idx >= y_idx:
        return slices, steps, None
    if y_idx == x_idx + 1:
        return slices, steps, None

    x = slices[x_idx]
    y = slices[y_idx]

    x_cats = [EXO_NAME] + x.categories_in_order
    y_cats = [EXO_NAME] + y.categories_in_order
    x_map = {c: i for i, c in enumerate(x_cats)}
    y_map = {c: j for j, c in enumerate(y_cats)}

    x_elem_to_cat = {e: c for c, es in x.cat_to_elems.items() for e in es}
    y_elem_to_cat = {e: c for c, es in y.cat_to_elems.items() for e in es}
    elems = set(x_elem_to_cat.keys()) | set(y_elem_to_cat.keys())

    Acol = np.zeros((len(x_cats), len(y_cats)), dtype=int)
    for e in elems:
        i = x_map.get(x_elem_to_cat.get(e, EXO_NAME), 0)
        j = y_map.get(y_elem_to_cat.get(e, EXO_NAME), 0)
        if i == 0 and j == 0:
            continue
        Acol[i, j] += 1

    col_step = StepModel(src_slice=x.slice_id, dst_slice=y.slice_id, src_cats=x_cats, dst_cats=y_cats, A=Acol)

    new_slices = slices[:x_idx + 1] + slices[y_idx:]
    new_steps: List[StepModel] = []
    new_steps.extend(steps[:x_idx])
    new_steps.append(col_step)
    new_steps.extend(steps[y_idx:])
    return new_slices, new_steps, col_step


def collapse_reorder(
        slices: List[SliceModel],
        steps: List[StepModel],
        col_step: StepModel,
        base_order: Dict[str, List[str]],
) -> Dict[str, List[str]]:
    """
    Single-pass reorder for the two collapsed anchor slices (PDF Eq. 6 & 7).

    After collapse, only θ_x and θ_y are reordered using each other's positions.
    No multi-iteration sweep is used. Other slices keep their base_order.

    Steps:
      1. Compute p_x(i) from θ_y positions (Eq. 6), reorder θ_x
      2. Compute p_y(j) from updated θ_x positions (Eq. 7), reorder θ_y
    """
    order = {sid: list(o) for sid, o in base_order.items()}

    # Find the two anchor slices
    x_sid = col_step.src_slice
    y_sid = col_step.dst_slice
    x_sl = None
    y_sl = None
    for sl in slices:
        if sl.slice_id == x_sid:
            x_sl = sl
        if sl.slice_id == y_sid:
            y_sl = sl
    if x_sl is None or y_sl is None:
        return order

    Acol = col_step.A
    x_cats = col_step.src_cats  # includes EXO at 0
    y_cats = col_step.dst_cats  # includes EXO at 0

    w_x = {c: float(max(1, len(x_sl.cat_to_elems.get(c, set())))) for c in x_sl.categories_in_order}
    w_y = {c: float(max(1, len(y_sl.cat_to_elems.get(c, set())))) for c in y_sl.categories_in_order}

    # Step 1: Compute p_x(i) from θ_y positions (Eq. 6), then reorder θ_x
    y_pos = {cat: p for p, cat in enumerate(order[y_sid])}
    pos0_x = {cat: p for p, cat in enumerate(x_sl.categories_in_order)}

    p_leaf_x: Dict[str, float] = {}
    for cat in x_sl.categories_in_order:
        if cat not in x_cats:
            p_leaf_x[cat] = float(pos0_x.get(cat, 0))
            continue
        i = x_cats.index(cat)
        denom = 0
        num = 0.0
        for j, dst_cat in enumerate(y_cats):
            if j == 0:
                continue
            v = int(Acol[i, j])
            if v <= 0:
                continue
            denom += v
            num += v * float(y_pos.get(dst_cat, 0.0))
        p_leaf_x[cat] = float(num / denom) if denom > 0 else float(pos0_x.get(cat, 0))

    order[x_sid] = interval_tree_reorder(x_sl.interval_tree, order[x_sid], p_leaf_x, w_x)

    # Step 2: Compute p_y(j) from updated θ_x positions (Eq. 7), then reorder θ_y
    x_pos = {cat: p for p, cat in enumerate(order[x_sid])}
    pos0_y = {cat: p for p, cat in enumerate(y_sl.categories_in_order)}

    p_leaf_y: Dict[str, float] = {}
    for cat in y_sl.categories_in_order:
        if cat not in y_cats:
            p_leaf_y[cat] = float(pos0_y.get(cat, 0))
            continue
        j = y_cats.index(cat)
        denom = 0
        num = 0.0
        for i, src_cat in enumerate(x_cats):
            if i == 0:
                continue
            v = int(Acol[i, j])
            if v <= 0:
                continue
            denom += v
            num += v * float(x_pos.get(src_cat, 0.0))
        p_leaf_y[cat] = float(num / denom) if denom > 0 else float(pos0_y.get(cat, 0))

    order[y_sid] = interval_tree_reorder(y_sl.interval_tree, order[y_sid], p_leaf_y, w_y)

    return order


# 7) Layout and drawing

@dataclass
class LayoutInfo:
    block_xy: Dict[BlockKey, Tuple[float, float]]
    block_bbox: Dict[BlockKey, Tuple[float, float, float, float]]  # x0,x1,y0,y1
    block_widths: Dict[BlockKey, float] = dataclasses.field(default_factory=dict)  # per-block width
    y_step: float = 0.0  # layer-to-layer center distance (includes block_height + layer_gap + vertical padding)


def _text_units_eaw(s: str) -> float:
    """
    统一的字符“宽度单位”估算：
    - 空格更窄
    - CJK (W/F) 更宽
    - 拉丁/数字中等
    """
    s = "" if s is None else str(s)
    units = 0.0
    for ch in s:
        if ch.isspace():
            units += 0.35
        elif unicodedata.east_asian_width(ch) in ("W", "F"):
            units += 1.00
        else:
            units += 0.60
    return max(1.0, units)


def estimate_label_width_data_units(text: str, cfg: StyleConfig) -> float:
    """
    把标签文本宽度估算成“数据坐标单位”，用于：
      - compute_layout: 预留组间距
      - apply_enclosure_label_visibility: 计算 group 标签左侧延伸范围
      - 以及 x-range 扩展估算（避免截断）
    """
    units = _text_units_eaw(text)
    # 用 block_width * ratio 作为“每单位字符”的数据坐标宽度，再加 padding
    char_w = float(cfg.block_width) * float(cfg.enclosure_label_mask_char_w_ratio)

    # 同时按字号比例做一次缩放（避免未来调大/调小 label 字号后估算失真）
    font_scale = float(cfg.enclosure_label_text_size) / float(cfg.block_text_size) if cfg.block_text_size else 1.0
    w = units * char_w * font_scale + 2.0 * float(cfg.enclosure_label_mask_pad_x)
    return float(w)


def compute_layout(slices: List[SliceModel], order: Dict[str, List[str]], cfg: StyleConfig,
                   block_mode: str = "existence", median_width: float = None) -> LayoutInfo:
    """
    Centering rule (to match your reference figure):
    - For EACH slice (layer), layout BODY blocks (non-Exo) in its own local x.
    - Then shift that slice's BODY so that its own center aligns to x=0 (per-layer centering).
    - Exo blocks are placed as a separate left column (NOT included in centering).
      Exo x is computed from the left-most BODY edge across ALL slices, so Exo stays aligned vertically.

    block_mode:
      - "existence": all blocks have the same width (cfg.block_width)
      - "strength": block width is proportional to element count, with median = median_width
    """
    block_xy: Dict[BlockKey, Tuple[float, float]] = {}
    block_bbox: Dict[BlockKey, Tuple[float, float, float, float]] = {}
    block_widths: Dict[BlockKey, float] = {}

    if median_width is None:
        median_width = cfg.block_width

    # Keep BODY keys per slice for per-layer shifting
    body_keys_by_slice: Dict[str, List[BlockKey]] = {}

    # Pre-compute element counts and block widths for strength mode
    if block_mode == "strength":
        # Collect all element counts
        all_elem_counts = []
        for sl in slices:
            for c in sl.categories_in_order:
                elem_count = len(sl.cat_to_elems.get(c, set()))
                if elem_count > 0:
                    all_elem_counts.append(elem_count)

        if all_elem_counts:
            # Compute median element count
            sorted_counts = sorted(all_elem_counts)
            n = len(sorted_counts)
            if n % 2 == 0:
                median_count = (sorted_counts[n // 2 - 1] + sorted_counts[n // 2]) / 2.0
            else:
                median_count = sorted_counts[n // 2]

            if median_count < 1:
                median_count = 1.0

            # Compute width for each block: width = (elem_count / median_count) * median_width
            for sl in slices:
                for c in sl.categories_in_order:
                    bk = BlockKey(sl.slice_id, c)
                    elem_count = len(sl.cat_to_elems.get(c, set()))
                    if elem_count < 1:
                        elem_count = 1
                    block_widths[bk] = (elem_count / median_count) * median_width
                # Exo block uses median width
                exo_bk = BlockKey(sl.slice_id, EXO_NAME)
                block_widths[exo_bk] = median_width
        else:
            # Fallback to uniform width
            for sl in slices:
                for c in sl.categories_in_order:
                    bk = BlockKey(sl.slice_id, c)
                    block_widths[bk] = median_width
                exo_bk = BlockKey(sl.slice_id, EXO_NAME)
                block_widths[exo_bk] = median_width
    else:
        # Existence mode: all blocks have the same width
        for sl in slices:
            for c in sl.categories_in_order:
                bk = BlockKey(sl.slice_id, c)
                block_widths[bk] = cfg.block_width
            exo_bk = BlockKey(sl.slice_id, EXO_NAME)
            block_widths[exo_bk] = cfg.block_width

    # 1) Layout BODY blocks per slice in a local coordinate system

    # --- Pre-compute enclosure padding inflation ---
    # Gap definitions:
    #   Enclosure gap = distance between enclosure BORDERS (not block edges)
    #   Layer gap = distance between outermost enclosure BORDERS of adjacent layers
    # The block gap in layout must be inflated by the total padding that enclosures
    # contribute from both sides, so that the final border-to-border gap matches the panel value.
    sg_styles = getattr(cfg, "supergroup_styles", {}) or {}

    def _sg_pad(level, key, default=0.0):
        """Get UI padding for a supergroup level."""
        s = sg_styles.get(level, {})
        try:
            v = s.get(key)
            return float(v) if v is not None and v != "" else default
        except Exception:
            return default

    # Horizontal inflation per boundary level:
    # For boundary at level k, total protrusion from BOTH sides into the gap:
    #   2*group_pad_x + Σ_{j=1}^{k} (2*step_x + pad_left_j + pad_right_j)
    max_agg_levels = max((len(sl.agg_cols) for sl in slices), default=0)
    _h_inflation: Dict[int, float] = {}  # boundary_level -> inflation
    for k in range(max_agg_levels):
        infl = 2.0 * cfg.group_enclosure_pad_x
        for j in range(1, k + 1):
            infl += 2.0 * cfg.enclosure_pad_step_x + _sg_pad(j, "pad_left") + _sg_pad(j, "pad_right")
        _h_inflation[k] = infl

    # Vertical inflation: total enclosure protrusion above and below blocks
    # top_pad = group_pad_y + Σ_{j=1}^{max_level} (step_y + pad_top_j)
    # bottom_pad = group_pad_y + Σ_{j=1}^{max_level} (step_y + pad_bottom_j)
    # Only add inflation when enclosures actually exist (max_agg_levels > 0)
    _max_sg_level = max(0, max_agg_levels - 1)  # number of supergroup levels
    if max_agg_levels > 0:
        _total_pad_top = cfg.group_enclosure_pad_y
        _total_pad_bottom = cfg.group_enclosure_pad_y
        for j in range(1, _max_sg_level + 1):
            _total_pad_top += cfg.enclosure_pad_step_y + _sg_pad(j, "pad_top")
            _total_pad_bottom += cfg.enclosure_pad_step_y + _sg_pad(j, "pad_bottom")
    else:
        _total_pad_top = 0.0
        _total_pad_bottom = 0.0

    # y spacing: layer_gap is the distance between outermost enclosure borders.
    # center-to-center = block_height + layer_gap + total_vertical_padding
    _y_step = cfg.block_height + cfg.layer_gap + _total_pad_top + _total_pad_bottom

    for t, sl in enumerate(slices):
        sid = sl.slice_id
        y_center = -t * _y_step

        cats = order.get(sid, sl.categories_in_order)

        inner_group_col = sl.agg_cols[0] if sl.agg_cols else None

        x = 0.0
        keys: List[BlockKey] = []

        prev_group_value: Optional[str] = None
        if inner_group_col is not None and len(cats) > 0:
            prev_group_value = sl.level_membership.get(inner_group_col, {}).get(cats[0], "")

        # Build previous membership values for all agg_cols levels
        prev_memberships: List[Optional[str]] = []
        for ac in sl.agg_cols:
            if len(cats) > 0:
                prev_memberships.append(sl.level_membership.get(ac, {}).get(cats[0], ""))
            else:
                prev_memberships.append(None)

        # Get enclosure gaps map (level_index -> gap)
        # Only the highest-level enclosure gap is used (to avoid redundancy/conflicts)
        enclosure_gaps = getattr(cfg, "enclosure_gaps", {})
        if enclosure_gaps:
            max_enc_level = max(enclosure_gaps.keys())
            highest_enc_gap = float(enclosure_gaps[max_enc_level])
        elif _max_sg_level > 0:
            # No explicit gaps set, use default for highest supergroup level
            highest_enc_gap = 2.0 if _max_sg_level == 1 else 3.0
        else:
            highest_enc_gap = 2.0

        for idx, c in enumerate(cats):
            bk = BlockKey(sid, c)
            bw = block_widths.get(bk, cfg.block_width)

            if idx == 0:
                x = bw / 2.0  # Start at half-width so left edge is at 0
            else:
                prev_bk = BlockKey(sid, cats[idx - 1])
                prev_bw = block_widths.get(prev_bk, cfg.block_width)

                if inner_group_col is None:
                    gap = cfg.outer_gap
                else:
                    # Check all agg_cols levels from outermost to innermost
                    # Use the gap of the outermost level where a boundary occurs
                    gap = cfg.inner_gap  # default: same group at all levels
                    boundary_level = -1
                    for lv_idx, ac in enumerate(sl.agg_cols):
                        cur_val = sl.level_membership.get(ac, {}).get(c, "")
                        prev_val = prev_memberships[lv_idx]
                        if cur_val != prev_val:
                            boundary_level = lv_idx
                        prev_memberships[lv_idx] = cur_val

                    if boundary_level >= 0:
                        if boundary_level == 0:
                            # Boundary at innermost aggregation level → outer_gap
                            gap = cfg.outer_gap + _h_inflation.get(0, 0.0)
                        else:
                            # Boundary at any enclosure level → use the single highest-level gap
                            gap = highest_enc_gap + _h_inflation.get(boundary_level, 0.0)

                    prev_group_value = sl.level_membership.get(inner_group_col, {}).get(c, "")

                x = x + prev_bw / 2.0 + gap + bw / 2.0

            keys.append(bk)
            block_xy[bk] = (x, y_center)
            block_bbox[bk] = (
                x - bw / 2, x + bw / 2,
                y_center - cfg.block_height / 2, y_center + cfg.block_height / 2
            )

        body_keys_by_slice[sid] = keys

    # 2) Per-layer centering: shift each slice's BODY to x=0 by its own bbox center
    for sl in slices:
        sid = sl.slice_id
        keys = body_keys_by_slice.get(sid, [])
        if not keys:
            continue

        x0s = [block_bbox[bk][0] for bk in keys]
        x1s = [block_bbox[bk][1] for bk in keys]
        layer_min_x0 = min(x0s)
        layer_max_x1 = max(x1s)
        layer_center = (layer_min_x0 + layer_max_x1) / 2.0

        shift = -layer_center  # make this slice centered at x=0

        for bk in keys:
            cx, cy = block_xy[bk]
            block_xy[bk] = (cx + shift, cy)
            x0, x1, y0, y1 = block_bbox[bk]
            block_bbox[bk] = (x0 + shift, x1 + shift, y0, y1)

    # 3) Compute global left-most BODY edge (after per-layer shifts) to place EXO column
    all_body_bboxes = []
    for sl in slices:
        sid = sl.slice_id
        for bk in body_keys_by_slice.get(sid, []):
            all_body_bboxes.append(block_bbox[bk])

    # Allow empty startup (no data loaded yet): return an empty layout without error.
    if not all_body_bboxes:
        return LayoutInfo(block_xy=block_xy, block_bbox=block_bbox, block_widths=block_widths, y_step=_y_step)

    global_body_left = min(bb[0] for bb in all_body_bboxes)

    # 4) Place Exo blocks as a stable left column (same x for all slices)
    # Use median_width for Exo in strength mode
    # NOTE: This is a placeholder position using body_left. The final position
    # is set after enclosures are built, using enclosure_global_x0 as reference.
    exo_width = median_width if block_mode == "strength" else cfg.block_width

    exo_cx = global_body_left - cfg.exo_gap - exo_width / 2
    for t, sl in enumerate(slices):
        sid = sl.slice_id
        y_center = -t * _y_step
        exo = BlockKey(sid, EXO_NAME)
        block_xy[exo] = (exo_cx, y_center)
        block_bbox[exo] = (
            exo_cx - exo_width / 2, exo_cx + exo_width / 2,
            y_center - cfg.block_height / 2, y_center + cfg.block_height / 2
        )
        block_widths[exo] = exo_width

    return LayoutInfo(block_xy=block_xy, block_bbox=block_bbox, block_widths=block_widths, y_step=_y_step)


def _bezier_points(p0, p1, p2, p3, n: int) -> List[Tuple[float, float]]:
    pts = []
    for k in range(n + 1):
        t = k / n
        x = (1 - t) ** 3 * p0[0] + 3 * (1 - t) ** 2 * t * p1[0] + 3 * (1 - t) * t ** 2 * p2[0] + t ** 3 * p3[0]
        y = (1 - t) ** 3 * p0[1] + 3 * (1 - t) ** 2 * t * p1[1] + 3 * (1 - t) * t ** 2 * p2[1] + t ** 3 * p3[1]
        pts.append((x, y))
    return pts


def band_polygon_straight_normal_flush(src_bbox, dst_bbox, thickness: float):
    """
    Straight band with constant NORMAL thickness.
    Endpoints are snapped to the source bottom edge (y_src) and destination top edge (y_dst)
    by moving along the band direction u, so there is NO gap at block connections.
    """
    x0s, x1s, y0s, y1s = src_bbox  # y0s = bottom
    x0d, x1d, y0d, y1d = dst_bbox  # y1d = top

    # Midpoints on the connecting edges
    p0 = ((x0s + x1s) / 2.0, y0s)  # source bottom midpoint
    p1 = ((x0d + x1d) / 2.0, y1d)  # dest top midpoint

    dx = p1[0] - p0[0]
    dy = p1[1] - p0[1]
    L = math.hypot(dx, dy)
    if L < 1e-9:
        # Degenerate: return a tiny vertical rectangle
        h = thickness / 2.0
        return [p0[0] - h, p0[0] + h, p0[0] + h, p0[0] - h, p0[0] - h], [p0[1], p0[1], p1[1], p1[1], p0[1]]

    ux, uy = dx / L, dy / L
    # unit normal
    nx, ny = -uy, ux
    h = thickness / 2.0

    def snap_to_y(pt, y_target):
        # Move along u so that y becomes y_target (preserves normal offset geometry)
        if abs(uy) < 1e-9:
            return (pt[0], y_target)
        t = (y_target - pt[1]) / uy
        return (pt[0] + t * ux, y_target)

    # Offset points at both ends
    sL = (p0[0] - h * nx, p0[1] - h * ny)
    sR = (p0[0] + h * nx, p0[1] + h * ny)
    dL = (p1[0] - h * nx, p1[1] - h * ny)
    dR = (p1[0] + h * nx, p1[1] + h * ny)

    # Snap endpoints back onto the block edge lines (this removes the notch/gap)
    sL = snap_to_y(sL, y0s)
    sR = snap_to_y(sR, y0s)
    dL = snap_to_y(dL, y1d)
    dR = snap_to_y(dR, y1d)

    xs = [sL[0], sR[0], dR[0], dL[0], sL[0]]
    ys = [sL[1], sR[1], dR[1], dL[1], sL[1]]
    return xs, ys


def allocate_ports_for_step(step: StepModel, layout: LayoutInfo, cfg: StyleConfig, proportion: float = 1.0,
                            exo_band_widths: Dict[tuple, float] = None):
    """
    Allocate port segments on block edges for bands.

    proportion: 0.0 ~ 1.0, controls how much of block width is used by bands.
                When < 1.0, gaps are evenly distributed between bands.
    exo_band_widths: Pre-computed widths for Exo-related bands (keyed by band key tuple).
                     If provided, these widths are used for both ends of Exo bands.
    """
    src_port = {}
    dst_port = {}

    src_order = sorted(step.src_cats, key=lambda c: layout.block_xy.get(BlockKey(step.src_slice, c), (0.0, 0.0))[0])
    dst_order = sorted(step.dst_cats, key=lambda c: layout.block_xy.get(BlockKey(step.dst_slice, c), (0.0, 0.0))[0])
    src_pos = {c: i for i, c in enumerate(src_order)}
    dst_pos = {c: i for i, c in enumerate(dst_order)}

    A = step.A
    proportion = max(0.1, min(1.0, float(proportion)))

    if exo_band_widths is None:
        exo_band_widths = {}

    # Source bottom allocation
    for i, src_cat in enumerate(step.src_cats):
        bkey = BlockKey(step.src_slice, src_cat)
        if bkey not in layout.block_bbox:
            continue
        x0, x1, y0, y1 = layout.block_bbox[bkey]
        dests = [j for j in range(A.shape[1]) if
                 int(A[i, j]) > 0 and not (src_cat == EXO_NAME and step.dst_cats[j] == EXO_NAME)]
        if not dests:
            continue
        dests.sort(key=lambda j: dst_pos.get(step.dst_cats[j], 0))

        if src_cat == EXO_NAME:
            # Exo as source (Inflow): use pre-computed widths
            # Collect widths for bands from this Exo
            band_widths = []
            for j in dests:
                band_key = (step.src_slice, src_cat, step.dst_slice, step.dst_cats[j])
                w = exo_band_widths.get(band_key, 0.0)
                band_widths.append((j, w))

            total_band_width = sum(w for _, w in band_widths)
            n_bands = len(band_widths)
            block_width = x1 - x0

            if n_bands > 0 and total_band_width > 0:
                # Center the bands on the Exo block
                total_gap = block_width - total_band_width
                gap_per = total_gap / (n_bands + 1) if total_gap > 0 else 0.0

                cur = x0 + gap_per
                for j, w in band_widths:
                    seg = (cur, cur + w)
                    cur = seg[1] + gap_per
                    src_port[(step.src_slice, src_cat, step.dst_slice, step.dst_cats[j])] = seg
        else:
            # Normal source block
            total = sum(int(A[i, j]) for j in dests)
            if total <= 0:
                continue

            block_width = x1 - x0
            n_bands = len(dests)

            # Calculate band widths and gaps
            total_band_width = block_width * proportion
            total_gap = block_width * (1.0 - proportion)

            # Distribute gaps evenly: n_bands + 1 gaps (including both ends)
            if n_bands > 0:
                gap_per = total_gap / (n_bands + 1)
            else:
                gap_per = 0.0

            scale = total_band_width / total

            cur = x0 + gap_per  # Start after first gap
            for j in dests:
                v = int(A[i, j])
                band_w = v * scale
                seg = (cur, cur + band_w)
                cur = seg[1] + gap_per  # Add gap after each band
                src_port[(step.src_slice, src_cat, step.dst_slice, step.dst_cats[j])] = seg

    # Destination top allocation
    for j, dst_cat in enumerate(step.dst_cats):
        bkey = BlockKey(step.dst_slice, dst_cat)
        if bkey not in layout.block_bbox:
            continue
        x0, x1, y0, y1 = layout.block_bbox[bkey]
        srcs = [i for i in range(A.shape[0]) if
                int(A[i, j]) > 0 and not (step.src_cats[i] == EXO_NAME and dst_cat == EXO_NAME)]
        if not srcs:
            continue
        srcs.sort(key=lambda i: src_pos.get(step.src_cats[i], 0))

        if dst_cat == EXO_NAME:
            # Exo as destination (Outflow): use pre-computed widths
            # Collect widths for bands to this Exo
            band_widths = []
            for i in srcs:
                band_key = (step.src_slice, step.src_cats[i], step.dst_slice, dst_cat)
                w = exo_band_widths.get(band_key, 0.0)
                band_widths.append((i, w))

            total_band_width = sum(w for _, w in band_widths)
            n_bands = len(band_widths)
            block_width = x1 - x0

            if n_bands > 0 and total_band_width > 0:
                # Center the bands on the Exo block
                total_gap = block_width - total_band_width
                gap_per = total_gap / (n_bands + 1) if total_gap > 0 else 0.0

                cur = x0 + gap_per
                for i, w in band_widths:
                    seg = (cur, cur + w)
                    cur = seg[1] + gap_per
                    dst_port[(step.src_slice, step.src_cats[i], step.dst_slice, dst_cat)] = seg
        else:
            # Normal destination block
            total = sum(int(A[i, j]) for i in srcs)
            if total <= 0:
                continue

            block_width = x1 - x0
            n_bands = len(srcs)

            # Calculate band widths and gaps
            total_band_width = block_width * proportion
            total_gap = block_width * (1.0 - proportion)

            # Distribute gaps evenly: n_bands + 1 gaps (including both ends)
            if n_bands > 0:
                gap_per = total_gap / (n_bands + 1)
            else:
                gap_per = 0.0

            scale = total_band_width / total

            cur = x0 + gap_per  # Start after first gap
            for i in srcs:
                v = int(A[i, j])
                band_w = v * scale
                seg = (cur, cur + band_w)
                cur = seg[1] + gap_per  # Add gap after each band
                dst_port[(step.src_slice, step.src_cats[i], step.dst_slice, dst_cat)] = seg

    return src_port, dst_port


def band_polygon_strength(src_seg, dst_seg, y_src, y_dst, cfg: StyleConfig):
    """
    Draw a band polygon for strength mode.

    src_seg: (x_left, x_right) on the source block's bottom edge (y = y_src)
    dst_seg: (x_left, x_right) on the destination block's top edge (y = y_dst)

    Returns xs, ys for a closed polygon (trapezoid shape).
    """
    x0s, x1s = src_seg  # source left, source right
    x0d, x1d = dst_seg  # dest left, dest right

    # Simple straight trapezoid: connect corners directly
    # Bottom edge: src_seg on y_src
    # Top edge: dst_seg on y_dst
    xs = [x0s, x1s, x1d, x0d, x0s]
    ys = [y_src, y_src, y_dst, y_dst, y_src]

    return xs, ys


def band_polygon_strength_curved(src_seg, dst_seg, y_src, y_dst, cfg: StyleConfig):
    """
    Draw a curved band polygon for strength mode (Bezier curves on left and right edges).

    src_seg: (x_left, x_right) on the source block's bottom edge (y = y_src)
    dst_seg: (x_left, x_right) on the destination block's top edge (y = y_dst)

    Returns xs, ys for a closed polygon with smooth curved sides.
    """
    x0s, x1s = src_seg  # source left, source right
    x0d, x1d = dst_seg  # dest left, dest right
    dy = y_dst - y_src
    c = cfg.band_curve_strength

    # Left edge Bezier control points
    p0L, p3L = (x0s, y_src), (x0d, y_dst)
    p1L, p2L = (x0s, y_src + c * dy), (x0d, y_dst - c * dy)

    # Right edge Bezier control points
    p0R, p3R = (x1s, y_src), (x1d, y_dst)
    p1R, p2R = (x1s, y_src + c * dy), (x1d, y_dst - c * dy)

    left = _bezier_points(p0L, p1L, p2L, p3L, cfg.band_samples)
    right = _bezier_points(p0R, p1R, p2R, p3R, cfg.band_samples)

    # Build polygon: bottom edge -> right curve -> top edge -> left curve (reversed)
    xs = [x0s, x1s]
    ys = [y_src, y_src]
    xs += [p[0] for p in right]
    ys += [p[1] for p in right]
    xs += [x1d, x0d]
    ys += [y_dst, y_dst]
    xs += [p[0] for p in reversed(left)]
    ys += [p[1] for p in reversed(left)]
    xs.append(xs[0])
    ys.append(ys[0])

    return xs, ys


def straight_band_polygon(p_src, p_dst, width: float):
    """
    width = distance between the two parallel edges (normal thickness) in DATA units.
    """
    x1, y1 = p_src
    x2, y2 = p_dst
    dx = x2 - x1
    dy = y2 - y1
    L = (dx * dx + dy * dy) ** 0.5
    if L < 1e-9:
        px, py = 0.0, 1.0
    else:
        px = -dy / L
        py = dx / L

    ox = px * (width / 2.0)
    oy = py * (width / 2.0)

    xs = [x1 + ox, x1 - ox, x2 - ox, x2 + ox, x1 + ox]
    ys = [y1 + oy, y1 - oy, y2 - oy, y2 + oy, y1 + oy]
    return xs, ys


def band_polygon(src_seg, dst_seg, y_src, y_dst, cfg: StyleConfig):
    x0s, x1s = src_seg
    x0d, x1d = dst_seg
    dy = y_dst - y_src
    c = cfg.band_curve_strength

    p0L, p3L = (x0s, y_src), (x0d, y_dst)
    p1L, p2L = (x0s, y_src + c * dy), (x0d, y_dst - c * dy)

    p0R, p3R = (x1s, y_src), (x1d, y_dst)
    p1R, p2R = (x1s, y_src + c * dy), (x1d, y_dst - c * dy)

    left = _bezier_points(p0L, p1L, p2L, p3L, cfg.band_samples)
    right = _bezier_points(p0R, p1R, p2R, p3R, cfg.band_samples)

    xs = [x0s, x1s]
    ys = [y_src, y_src]
    xs += [p[0] for p in right]
    ys += [p[1] for p in right]
    xs += [x1d, x0d]
    ys += [y_dst, y_dst]
    xs += [p[0] for p in reversed(left)]
    ys += [p[1] for p in reversed(left)]
    xs.append(xs[0]);
    ys.append(ys[0])
    return xs, ys


def enforce_sibling_gap_x(boxes: List[dict], gap_x: float) -> List[dict]:
    """
    boxes: list of dict with keys x0,x1,y0,y1,... for ONE slice + ONE level.
    Enforce minimal horizontal gap between adjacent boxes by shrinking them symmetrically at the boundary.
    """
    if len(boxes) <= 1:
        return boxes

    boxes = sorted(boxes, key=lambda b: b["x0"])

    for k in range(len(boxes) - 1):
        a = boxes[k]
        b = boxes[k + 1]

        cur_gap = b["x0"] - a["x1"]
        if cur_gap >= gap_x:
            continue

        mid = (a["x1"] + b["x0"]) / 2.0
        a["x1"] = mid - gap_x / 2.0
        b["x0"] = mid + gap_x / 2.0

        # safety clamp: avoid inverted boxes
        if a["x1"] <= a["x0"]:
            a["x1"] = a["x0"] + 1e-6
        if b["x1"] <= b["x0"]:
            b["x1"] = b["x0"] + 1e-6

    return boxes


import plotly.graph_objects as go


def add_enclosure_name_overlay(
        fig,
        *,
        x0: float,
        x1: float,
        y_top: float,
        text: str,
        font_size: int,
        y_offset: float,
        pad_x: float,
        pad_y: float,
        bgcolor: str = "white",
):
    """
    在围框顶部“覆盖绘制”名字，不改变围框几何尺寸。
    - y_top: 围框的真实上边界（不要为了名字上移）
    - y_offset: 名字相对 y_top 向上的偏移（数据坐标单位）
    - pad_x/pad_y: 白底 mask 的 padding（数据坐标单位）
    """
    if (text is None) or (str(text).strip() == ""):
        return

    text = str(text)

    xc = (x0 + x1) / 2.0
    yt = y_top - y_offset

    # 粗略估算文字宽度，避免 mask 过大：按字符数线性估计
    est_half_w = max(0.0, (len(text) * pad_x) / 2.0)

    # 白底 mask（盖住后面的线/框，避免“被盖住”）
    fig.add_shape(
        type="rect",
        x0=xc - est_half_w - pad_x,
        x1=xc + est_half_w + pad_x,
        y0=yt - pad_y,
        y1=yt + pad_y,
        line_width=0,
        fillcolor=bgcolor,
        layer="above",
    )

    # 文字（覆盖在最上层）
    fig.add_trace(
        go.Scatter(
            x=[xc],
            y=[yt],
            mode="text",
            text=[text],
            textposition="middle center",
            textfont=dict(size=font_size),
            hoverinfo="skip",
            showlegend=False,
            cliponaxis=False,
        )
    )


def build_figure(
        slices: List[SliceModel],
        steps: List[StepModel],
        order: Dict[str, List[str]],
        block_labels: Dict[BlockKey, Set[str]],
        band_labels: Dict[BandKey, Set[str]],
        cfg: StyleConfig,
        band_mode: str = "existence",
        band_proportion: float = 0.5,
        block_mode: str = "existence",
        block_median_width: float = None,
        collapse_state: dict = None,
        graph_size: Optional[dict] = None,
        initial_y_range: Optional[list] = None,
) -> Tuple[go.Figure, Dict[str, Any]]:
    if block_median_width is None:
        block_median_width = cfg.block_width

    # Resolve actual canvas pixel dimensions for text-wrapping calculations.
    # graph_size comes from the browser's measured Plotly div dimensions.
    # Fall back to conservative defaults when not yet known.
    _RIGHT_PANEL_W = 920
    _TOPBAR_H = 90
    _gs_w = (graph_size or {}).get("w") if graph_size else None
    _gs_h = (graph_size or {}).get("h") if graph_size else None
    _CANVAS_W = float(_gs_w) if (_gs_w and _gs_w > 100) else 600.0
    _CANVAS_H = float(_gs_h) if (_gs_h and _gs_h > 100) else 700.0

    layout = compute_layout(slices, order, cfg, block_mode=block_mode, median_width=block_median_width)

    def block_id(bk: BlockKey) -> str:
        return f"BLOCK::{bk.slice_id}::{bk.cat_name}"

    def band_id(bk: BandKey) -> str:
        return f"BAND::{bk.src_slice}::{bk.src_cat}::{bk.dst_slice}::{bk.dst_cat}"

    def disp_cat(c: str) -> str:
        return "Exo" if c == EXO_NAME else c

    block_to_bands: Dict[str, Set[str]] = defaultdict(set)
    band_to_blocks: Dict[str, Set[str]] = defaultdict(set)
    curve_to_id: List[str] = []

    fig = go.Figure()

    # Empty startup (no data loaded yet): return a blank canvas.
    # The "No data loaded" hint is shown by #no-data-hint-html (HTML div, JS-controlled).
    body_bboxes = [bb for bk, bb in layout.block_bbox.items() if bk.cat_name != EXO_NAME]
    if not body_bboxes:
        fig.update_xaxes(visible=False)
        fig.update_yaxes(visible=False)
        fig.update_layout(
            margin=dict(l=5, r=5, t=5, b=5),
            plot_bgcolor="white",
            paper_bgcolor="white",
        )
        meta: Dict[str, Any] = {"slices": [sl.slice_id for sl in slices]}
        meta["slice_label_anno_indices"] = {}
        return fig, meta

    # --- Pre-compute Exo block widths for strength mode ---
    # When band mode is strength, calculate Exo widths
    # based on the required space for Outflow/Inflow bands
    exo_band_widths: Dict[tuple, float] = {}

    if band_mode == "strength":
        # Step 1: Calculate widths for Exo-related bands (based on non-Exo end)
        exo_top_widths: Dict[BlockKey, float] = defaultdict(float)  # Exo as dest (Outflow)
        exo_bottom_widths: Dict[BlockKey, float] = defaultdict(float)  # Exo as src (Inflow)

        for st in steps:
            A = st.A
            src_order = sorted(st.src_cats, key=lambda c: layout.block_xy.get(BlockKey(st.src_slice, c), (0.0, 0.0))[0])
            dst_order = sorted(st.dst_cats, key=lambda c: layout.block_xy.get(BlockKey(st.dst_slice, c), (0.0, 0.0))[0])

            # Process Outflow bands (normal block -> Exo): width determined by source block
            for i, src_cat in enumerate(st.src_cats):
                if src_cat == EXO_NAME:
                    continue
                bkey = BlockKey(st.src_slice, src_cat)
                if bkey not in layout.block_bbox:
                    continue
                x0, x1 = layout.block_bbox[bkey][0], layout.block_bbox[bkey][1]
                block_width = x1 - x0

                # Find Exo destination
                exo_j = None
                for j, dst_cat in enumerate(st.dst_cats):
                    if dst_cat == EXO_NAME and int(A[i, j]) > 0:
                        exo_j = j
                        break

                if exo_j is not None:
                    # Calculate width at source end
                    dests = [j for j in range(A.shape[1]) if
                             int(A[i, j]) > 0 and not (src_cat == EXO_NAME and st.dst_cats[j] == EXO_NAME)]
                    total = sum(int(A[i, j]) for j in dests)
                    if total > 0:
                        total_band_width = block_width * band_proportion
                        scale = total_band_width / total
                        band_w = int(A[i, exo_j]) * scale

                        band_key = (st.src_slice, src_cat, st.dst_slice, EXO_NAME)
                        exo_band_widths[band_key] = band_w

                        exo_bk = BlockKey(st.dst_slice, EXO_NAME)
                        exo_top_widths[exo_bk] += band_w

            # Process Inflow bands (Exo -> normal block): width determined by dest block
            for j, dst_cat in enumerate(st.dst_cats):
                if dst_cat == EXO_NAME:
                    continue
                bkey = BlockKey(st.dst_slice, dst_cat)
                if bkey not in layout.block_bbox:
                    continue
                x0, x1 = layout.block_bbox[bkey][0], layout.block_bbox[bkey][1]
                block_width = x1 - x0

                # Find Exo source
                exo_i = None
                for i, src_cat in enumerate(st.src_cats):
                    if src_cat == EXO_NAME and int(A[i, j]) > 0:
                        exo_i = i
                        break

                if exo_i is not None:
                    # Calculate width at dest end
                    srcs = [i for i in range(A.shape[0]) if
                            int(A[i, j]) > 0 and not (st.src_cats[i] == EXO_NAME and dst_cat == EXO_NAME)]
                    total = sum(int(A[i, j]) for i in srcs)
                    if total > 0:
                        total_band_width = block_width * band_proportion
                        scale = total_band_width / total
                        band_w = int(A[exo_i, j]) * scale

                        band_key = (st.src_slice, EXO_NAME, st.dst_slice, dst_cat)
                        exo_band_widths[band_key] = band_w

                        exo_bk = BlockKey(st.src_slice, EXO_NAME)
                        exo_bottom_widths[exo_bk] += band_w

        # Step 2: Find the maximum required Exo width (unified for all Exo blocks)
        max_exo_width = 0.0
        for exo_bk in set(exo_top_widths.keys()) | set(exo_bottom_widths.keys()):
            top_w = exo_top_widths.get(exo_bk, 0.0)
            bottom_w = exo_bottom_widths.get(exo_bk, 0.0)
            required = max(top_w, bottom_w)
            if required > max_exo_width:
                max_exo_width = required

        # Add margin for gaps (proportion of total)
        if max_exo_width > 0:
            max_exo_width = max_exo_width / band_proportion

        # Ensure minimum Exo width
        min_exo_width = block_median_width if block_median_width else cfg.block_width
        max_exo_width = max(max_exo_width, min_exo_width)

        # Step 3: Update all Exo block bboxes in layout
        # When band_mode is strength, Exo blocks need to be wide enough to accommodate all bands
        for sl in slices:
            exo_bk = BlockKey(sl.slice_id, EXO_NAME)
            if exo_bk in layout.block_bbox:
                old_bbox = layout.block_bbox[exo_bk]
                cx, cy = layout.block_xy[exo_bk]
                new_half_w = max_exo_width / 2.0
                layout.block_bbox[exo_bk] = (
                    cx - new_half_w, cx + new_half_w,
                    old_bbox[2], old_bbox[3]
                )
                layout.block_widths[exo_bk] = max_exo_width

    # --- Precompute viewport ranges ONCE (also used by BG click-catcher) ---
    all_bboxes = list(layout.block_bbox.values())
    y_min = min(bb[2] for bb in all_bboxes) - cfg.viewport_pad_y
    y_max = max(bb[3] for bb in all_bboxes) + cfg.viewport_pad_y

    body_min_x0 = min(bb[0] for bb in body_bboxes)
    body_max_x1 = max(bb[1] for bb in body_bboxes)
    half_body = max(abs(body_min_x0), abs(body_max_x1))

    exo_bboxes = [layout.block_bbox[BlockKey(sl.slice_id, EXO_NAME)] for sl in slices]
    exo_min_x0 = min(bb[0] for bb in exo_bboxes)
    half_need_for_exo = max(0.0, -exo_min_x0)

    # Use actual content bounds instead of forcing symmetric range around x=0.
    # Left edge: exo block left edge (negative) with padding.
    # Right edge: slice labels sit to the right of body blocks.
    slice_label_right = body_max_x1 + 1.0 + cfg.block_width * 0.5 + cfg.viewport_pad_x
    x_left = min(exo_min_x0, body_min_x0) - cfg.viewport_pad_x
    x_right = max(body_max_x1, slice_label_right) + cfg.viewport_pad_x
    x_range = [x_left, x_right]
    y_range = initial_y_range if initial_y_range and len(initial_y_range) == 2 else [y_min, y_max]
    # A full-viewport filled polygon that is clickable on FILLS.
    fig.add_trace(go.Scatter(
        x=[x_range[0], x_range[1], x_range[1], x_range[0], x_range[0]],
        y=[y_range[0], y_range[0], y_range[1], y_range[1], y_range[0]],
        mode="lines",
        name="",
        line=dict(color="rgba(0,0,0,0)", width=0),
        fill="toself",
        fillcolor="rgba(0,0,0,0)",
        hoveron="fills",
        hoverinfo="text",
        text=[" ", " ", " ", " ", " "],
        hovertemplate="<extra></extra>",
        showlegend=False,
        opacity=1.0,
    ))
    curve_to_id.append("BG::CLICK")

    # --- Enclosures + their labels (labels as TRACES, below bands) ---
    shapes: List[dict] = []
    label_specs_all: List[dict] = []
    label_level_keys_all: List[str] = []
    mask_shape_indices: List[int] = []
    border_shape_indices: List[int] = []
    label_depths: List[int] = []

    outermost_enclosure_bboxes: List[Tuple[float, float, float, float]] = []

    # Effective ppu at the initial "contain-fit" display (zoom ≈ 0.88).
    # Using this for text truncation gives correct sizing at the canonical
    # initial view — the user's "treat initial state as a screenshot slice at
    # the contain-fit zoom level" principle.
    _CONTAIN_FIT_ZOOM = 0.88
    _eff_x_span = max(x_range[1] - x_range[0], 0.01)
    _eff_y_span = max(y_range[1] - y_range[0], 0.01)
    _eff_ppu_x = _CANVAS_W * _CONTAIN_FIT_ZOOM / _eff_x_span
    _eff_ppu_y = _CANVAS_H * _CONTAIN_FIT_ZOOM / _eff_y_span

    # ── Deterministic ppu for TEXT-FIT use only ────────────────────────────
    # The problem: graph_size (w, h) comes from the browser. Even with the
    # client-side 25px quantization, different browsers / window sizes still
    # give different _eff_ppu_x/y → different max_chars → different wrap →
    # visually different diagrams across browsers for identical data.
    #
    # Fix: cap the canvas dimensions used for text-wrapping at a fixed
    # reference (_TEXT_FIT_REF_W × _TEXT_FIT_REF_H). Any canvas larger than
    # the reference uses the reference — so on every "reasonably sized"
    # desktop browser, the text-fit sees the same inputs and produces the
    # same wrap. Smaller canvases still use their actual size so text
    # doesn't overflow on small screens.
    #
    # We keep the full _eff_ppu_x/y for other uses (enclosure layout, etc.)
    # because those affect geometry, not just text, and forcing a reference
    # size there would distort the layout.
    _TEXT_FIT_REF_W = 1600.0
    _TEXT_FIT_REF_H = 900.0
    _fit_cw = min(float(_CANVAS_W), _TEXT_FIT_REF_W)
    _fit_ch = min(float(_CANVAS_H), _TEXT_FIT_REF_H)
    _fit_ppu_x = _fit_cw * _CONTAIN_FIT_ZOOM / _eff_x_span
    _fit_ppu_y = _fit_ch * _CONTAIN_FIT_ZOOM / _eff_y_span

    for sl in slices:
        sid = sl.slice_id
        cats = order.get(sid, sl.categories_in_order)

        shape_offset = len(shapes)
        sh, ls, lk, mi, bi, ld, outermost_bb = build_enclosure_shapes_for_slice(
            sl, cats, layout, cfg, ppu_x=_eff_ppu_x)
        shapes.extend(sh)

        if outermost_bb is not None:
            outermost_enclosure_bboxes.append(outermost_bb)

        # shift shape indices to global
        mi_g = [shape_offset + int(x) for x in mi]
        bi_g = [shape_offset + int(x) for x in bi]
        mask_shape_indices.extend(mi_g)
        border_shape_indices.extend(bi_g)

        label_specs_all.extend(ls)
        label_level_keys_all.extend(lk)
        label_depths.extend(ld)

    # Compute global enclosure extent (widest enclosure left/right)
    if outermost_enclosure_bboxes:
        enclosure_global_x0 = min(bb[0] for bb in outermost_enclosure_bboxes)
        enclosure_global_x1 = max(bb[1] for bb in outermost_enclosure_bboxes)
    else:
        # Fallback to body extent if no enclosures
        enclosure_global_x0 = body_min_x0
        enclosure_global_x1 = body_max_x1

    # --- Reposition Exo blocks: Distance = gap from Exo right edge to outermost enclosure left ---
    # Use the wider of block_median_width and the band-calculated width from layout.block_widths
    _exo_base_width = block_median_width if block_mode == "strength" else cfg.block_width
    # Find the maximum Exo width across all slices (may have been enlarged by band strength calc)
    _exo_max_w = _exo_base_width
    for sl in slices:
        exo_bk = BlockKey(sl.slice_id, EXO_NAME)
        if exo_bk in layout.block_widths:
            _exo_max_w = max(_exo_max_w, layout.block_widths[exo_bk])
    exo_width = _exo_max_w
    new_exo_cx = enclosure_global_x0 - cfg.exo_gap - exo_width / 2.0
    for sl in slices:
        sid = sl.slice_id
        exo_bk = BlockKey(sid, EXO_NAME)
        if exo_bk in layout.block_xy:
            old_cx, cy = layout.block_xy[exo_bk]
            layout.block_xy[exo_bk] = (new_exo_cx, cy)
            layout.block_bbox[exo_bk] = (
                new_exo_cx - exo_width / 2.0, new_exo_cx + exo_width / 2.0,
                cy - cfg.block_height / 2.0, cy + cfg.block_height / 2.0
            )
            layout.block_widths[exo_bk] = exo_width
    cfg.exo_anchor_x = new_exo_cx

    # --- Recalculate viewport after repositioning ---
    all_bboxes = list(layout.block_bbox.values())
    y_min = min(bb[2] for bb in all_bboxes) - cfg.viewport_pad_y
    y_max = max(bb[3] for bb in all_bboxes) + cfg.viewport_pad_y
    exo_bboxes = [layout.block_bbox[BlockKey(sl.slice_id, EXO_NAME)] for sl in slices]
    exo_min_x0 = min(bb[0] for bb in exo_bboxes)
    slice_label_right = enclosure_global_x1 + 1.0 + cfg.viewport_pad_x
    x_left = min(exo_min_x0, body_min_x0) - cfg.viewport_pad_x
    x_right = max(body_max_x1, slice_label_right) + cfg.viewport_pad_x
    x_range = [x_left, x_right]
    y_range = initial_y_range if initial_y_range and len(initial_y_range) == 2 else [y_min, y_max]
    for idx, tr in enumerate(fig.data):
        if idx < len(curve_to_id) and curve_to_id[idx] == "BG::CLICK":
            tr.x = [x_range[0], x_range[1], x_range[1], x_range[0], x_range[0]]
            tr.y = [y_range[0], y_range[0], y_range[1], y_range[1], y_range[0]]
            break
    fig.update_xaxes(range=x_range)
    fig.update_yaxes(range=y_range, scaleanchor="x", scaleratio=1)

    # Attach shapes now (below all traces)
    fig.update_layout(
        shapes=shapes,
        margin=dict(l=5, r=5, t=5, b=5),
        plot_bgcolor="white",
        paper_bgcolor="white",
        hovermode="closest",
        hoverdistance=80,
        dragmode="pan",
        clickmode="event+select",
        font=dict(size=cfg.ui_font_size),
        hoverlabel=dict(
            bgcolor="#2b2b2b",
            bordercolor="#2b2b2b",
            # Pin font family explicitly so plotly's native block/band tooltip
            # uses the SAME font as our custom HTML tooltip (#_lbl_tt) for
            # enclosure / aggregation names. Without an explicit family,
            # plotly picks its default font stack which varies by browser
            # and renders glyphs at slightly different widths than Arial,
            # making the two tooltip styles look subtly different. With this
            # pin, the visual sizing is identical across Chrome / Firefox /
            # Safari / Edge.
            font=dict(size=cfg.tooltip_font_size, color="#ffffff",
                      family="Arial, sans-serif"),
            namelength=0,
            align="auto",
        ),
        # Allow the figure to be responsive but set minimum dimensions
        autosize=True,
        minreducedheight=600,
        minreducedwidth=700,
        # Keep zoom/pan state across Dash figure updates.
        # As long as uirevision stays the same constant value, Plotly
        # will never reset the axis ranges when the figure is replaced.
        uirevision="birdcage-fixed",
    )

    # Add enclosure labels:
    # - Supergroup labels as TRACES (must be BEFORE bands so bands draw above them).
    # - Group labels as ANNOTATIONS (always on top of everything, including shapes with layer="above").
    label_trace_indices: List[int] = [-1] * len(label_specs_all)
    group_anno_indices: List[int] = [-1] * len(label_specs_all)

    # 2) All enclosure labels as ANNOTATIONS (always above bands)
    annos: List[dict] = list(getattr(fig.layout, "annotations", []) or [])
    label_anno_indices: List[int] = [-1] * len(label_specs_all)

    for i, spec in enumerate(label_specs_all):
        k = str(label_level_keys_all[i])

        if k == "group":
            # left-side label (final x/y will be set in apply_enclosure_label_visibility)
            ann = dict(
                x=float(spec.get("x", 0.0)),
                y=float(spec.get("y", 0.0)),
                xref="x", yref="y",
                text=str(spec.get("text", "")),
                showarrow=False,
                visible=False,
                font=dict(size=cfg.enclosure_label_text_size, color=cfg.enclosure_label_text_color),
                bgcolor="rgba(0,0,0,0)",
                bordercolor="rgba(0,0,0,0)",
                borderpad=1,
                xanchor="right",
                yanchor="middle",
                align="right",
            )
        else:
            # top-center label (supergroup...) -> text only, NO background
            # White background is handled by separate shape (layer="below")
            ann = dict(
                x=float(spec.get("x", 0.0)),
                y=float(spec.get("y", 0.0)),
                xref="x", yref="y",
                text=str(spec.get("text", "")),
                showarrow=False,
                visible=False,
                font=dict(size=cfg.enclosure_label_text_size, color=cfg.enclosure_label_text_color),
                bgcolor="rgba(0,0,0,0)",  # Transparent background for supergroup labels
                bordercolor="rgba(0,0,0,0)",
                borderpad=1,
                xanchor="center",
                yanchor="bottom",  # bottom of the label sits at y (distance is measured from border to label bottom)
                align="center",
            )

        annos.append(ann)
        label_anno_indices[i] = int(len(annos) - 1)

    fig.update_layout(annotations=annos)

    # Band thickness/opacity are computed per band using cfg.band_width_ratio_by_type / cfg.band_opacity_by_type

    # Band z-order (BOTTOM -> TOP). Plotly draws later traces on top.
    BAND_Z = {
        "Unknown": -1,
        "Inheri": 0,
        "Merge": 1,
        "Split": 2,
        "SpMe": 3,
        "Inflow": 4,
        "Outflow": 5,
    }

    band_items: List[Dict[str, Any]] = []
    block_items: List[Dict[str, Any]] = []
    block_text_anno_indices: Dict[str, int] = {}  # Maps block_id to annotation index
    seq = 0

    # Pre-compute port allocations for strength mode
    all_src_ports: Dict[tuple, tuple] = {}
    all_dst_ports: Dict[tuple, tuple] = {}

    if band_mode == "strength":
        # Use exo_band_widths computed earlier (if in strength mode for both block and band)
        for st in steps:
            src_port, dst_port = allocate_ports_for_step(st, layout, cfg, proportion=band_proportion,
                                                         exo_band_widths=exo_band_widths)
            all_src_ports.update(src_port)
            all_dst_ports.update(dst_port)

    for st in steps:
        A = st.A
        for i, src_cat in enumerate(st.src_cats):
            for j, dst_cat in enumerate(st.dst_cats):
                v = int(A[i, j])
                if v <= 0 or (src_cat == EXO_NAME and dst_cat == EXO_NAME):
                    continue

                bsrc = layout.block_bbox[BlockKey(st.src_slice, src_cat)]
                bdst = layout.block_bbox[BlockKey(st.dst_slice, dst_cat)]

                bk = BandKey(st.src_slice, src_cat, st.dst_slice, dst_cat)

                labels_all = band_labels.get(bk, {"Unknown"})
                labels_evt = set(labels_all) & _BAND_EVENT_SET
                if not labels_evt:
                    labels_evt = {"Unknown"}

                event_text = format_band_event_text(labels_evt)
                primary = _primary_band_label(labels_evt)

                # Per-type geometry + opacity (fallback to global defaults)
                bw_map = getattr(cfg, "band_width_ratio_by_type", {}) or {}
                bo_map = getattr(cfg, "band_opacity_by_type", {}) or {}

                try:
                    width_ratio_i = float(bw_map.get(str(primary), float(cfg.band_width_ratio)))
                except Exception:
                    width_ratio_i = float(cfg.band_width_ratio)
                width_ratio_i = float(max(0.02, min(0.30, width_ratio_i)))
                thickness_i = float(cfg.block_width * width_ratio_i)

                try:
                    opacity_i = float(bo_map.get(str(primary), float(cfg.band_default_opacity)))
                except Exception:
                    opacity_i = float(cfg.band_default_opacity)
                opacity_i = float(max(0.0, min(1.0, opacity_i)))

                # Choose band polygon based on mode
                if band_mode == "strength":
                    # Strength mode: use allocated port segments
                    port_key = (st.src_slice, src_cat, st.dst_slice, dst_cat)
                    src_seg = all_src_ports.get(port_key)
                    dst_seg = all_dst_ports.get(port_key)

                    if src_seg is not None and dst_seg is not None:
                        # Get y coordinates from bboxes
                        y_src = bsrc[2]  # y0 = bottom of source block
                        y_dst = bdst[3]  # y1 = top of dest block
                        xs, ys = band_polygon_strength_curved(src_seg, dst_seg, y_src, y_dst, cfg)
                    else:
                        # Fallback to existence mode if port allocation failed
                        xs, ys = band_polygon_straight_normal_flush(bsrc, bdst, thickness_i)
                else:
                    # Existence mode: original behavior
                    xs, ys = band_polygon_straight_normal_flush(bsrc, bdst, thickness_i)

                hover_html_band = f"<span style='font-weight:normal'>{event_text}</span>"

                color = cfg.band_colors.get(primary, cfg.band_colors["Unknown"])
                bid = band_id(bk)

                src_block_oid = block_id(BlockKey(st.src_slice, src_cat))
                dst_block_oid = block_id(BlockKey(st.dst_slice, dst_cat))
                block_to_bands[src_block_oid].add(bid)
                block_to_bands[dst_block_oid].add(bid)
                band_to_blocks[bid].add(src_block_oid)
                band_to_blocks[bid].add(dst_block_oid)

                band_items.append({
                    "z": int(BAND_Z.get(primary, BAND_Z["Unknown"])),
                    "seq": int(seq),
                    "bid": bid,
                    "primary": str(primary),
                    "opacity": float(opacity_i),
                    "xs": xs,
                    "ys": ys,
                    "color": color,
                    "hover": hover_html_band,
                    "bsrc": bsrc,
                    "bdst": bdst,
                })
                seq += 1

    band_items.sort(key=lambda d: (d["z"], d["seq"]))

    def _poly_to_path(xs_in: List[float], ys_in: List[float]) -> str:
        """Convert a closed polygon into a Plotly SVG path string."""
        if not xs_in or not ys_in or len(xs_in) != len(ys_in):
            return ""
        pts = list(zip(xs_in, ys_in))
        if len(pts) < 3:
            return ""
        cmd = [f"M {pts[0][0]},{pts[0][1]}"]
        for x, y in pts[1:]:
            cmd.append(f"L {x},{y}")
        cmd.append("Z")
        return " ".join(cmd)

    # Visible bands are drawn as SHAPES so we can place them ABOVE group borders
    # (which are shapes with layer='above') but still BELOW supergroup label text
    # (which is drawn as annotations).
    # IMPORTANT: because shapes are not part of fig.data, we must also record
    # a mapping from BAND oid -> shape index so highlight logic can still
    # dim/strengthen connected bands when a block (or a band hover-catcher) is clicked.
    band_shapes: List[dict] = []
    band_shape_oids: List[str] = []
    for it in band_items:
        p = _poly_to_path(it["xs"], it["ys"])
        if not p:
            continue
        band_shapes.append(dict(
            type="path",
            path=p,
            line=dict(color=it["color"], width=0.5),
            fillcolor=it["color"],
            opacity=float(it.get("opacity", float(cfg.band_default_opacity))),
            layer="below",
        ))
        band_shape_oids.append(it["bid"])

    band_shape_indices: Dict[str, int] = {}
    if len(band_shapes) > 0:
        existing_shapes = list(getattr(fig.layout, "shapes", []) or [])
        start_idx = len(existing_shapes)
        fig.update_layout(shapes=existing_shapes + band_shapes)

        # record indices for highlight
        for k in range(len(band_shape_oids)):
            band_shape_indices[str(band_shape_oids[k])] = int(start_idx + k)

    # --- Band hover catchers (keep as traces for tooltip placement) ---
    for it in band_items:
        xs = it["xs"]
        ys = it["ys"]
        color = it["color"]
        hover_html_band = it["hover"]
        bid = it["bid"]
        bsrc = it["bsrc"]
        bdst = it["bdst"]

        # hover-catcher points along band centerline (tooltip near mouse position)
        x0s, x1s, y0s, y1s = bsrc
        x0d, x1d, y0d, y1d = bdst
        p0x, p0y = (x0s + x1s) / 2.0, y0s
        p1x, p1y = (x0d + x1d) / 2.0, y1d

        n_hover = 41
        hx = [p0x + (p1x - p0x) * (k / (n_hover - 1)) for k in range(n_hover)]
        hy = [p0y + (p1y - p0y) * (k / (n_hover - 1)) for k in range(n_hover)]

        fig.add_trace(go.Scatter(
            x=hx, y=hy,
            mode="markers",
            name="",
            marker=dict(
                size=18,
                color="rgba(0,0,0,0)",
                line=dict(width=0, color="rgba(0,0,0,0)")
            ),
            hoverinfo="text",
            text=[hover_html_band] * n_hover,
            hovertemplate="%{text}<extra></extra>",
            hoverlabel=dict(bgcolor=color, bordercolor=color),
            showlegend=False,
            opacity=1.0,
        ))
        curve_to_id.append(bid)

    # Blocks + labels

    # --- Text wrapping helpers (defined once, used per-block) ---
    def _is_cjk(ch):
        """Return True if character is CJK (approximately double-width)."""
        cp = ord(ch)
        return (0x4E00 <= cp <= 0x9FFF or 0x3400 <= cp <= 0x4DBF or
                0xF900 <= cp <= 0xFAFF or 0x20000 <= cp <= 0x2FA1F or
                0x3000 <= cp <= 0x303F or 0xFF00 <= cp <= 0xFFEF)

    def _visual_len(s):
        """Approximate visual length: CJK chars count as ~1.7 Latin chars."""
        return sum(1.7 if _is_cjk(ch) else 1.0 for ch in s)

    # Arial font metrics in em units (1 em = font_size px). Source: Adobe
    # Arial AFM (PostScript Font Metrics). Used by _wrap_text below to
    # make pixel-accurate line-break decisions instead of relying on a
    # single average-width estimate (_CHAR_W) — which fails the moment
    # the actual text deviates from the assumed average composition.
    # Wide letters (M, W, etc.) push real line width well past what the
    # average estimate predicted, and lines visibly overflow the block.
    # With this table, _wrap_text walks each character's TRUE width and
    # breaks the moment cumulative width hits the block edge — exactly
    # the "wall-bump → break" behaviour the user described.
    _ARIAL_EM = {
        ' ': 0.278, '!': 0.278, '"': 0.355, '#': 0.556, '$': 0.556, '%': 0.889,
        '&': 0.667, "'": 0.191, '(': 0.333, ')': 0.333, '*': 0.389, '+': 0.584,
        ',': 0.278, '-': 0.333, '.': 0.278, '/': 0.278,
        '0': 0.556, '1': 0.556, '2': 0.556, '3': 0.556, '4': 0.556,
        '5': 0.556, '6': 0.556, '7': 0.556, '8': 0.556, '9': 0.556,
        ':': 0.278, ';': 0.278, '<': 0.584, '=': 0.584, '>': 0.584, '?': 0.556,
        '@': 1.015,
        'A': 0.667, 'B': 0.667, 'C': 0.722, 'D': 0.722, 'E': 0.667, 'F': 0.611,
        'G': 0.778, 'H': 0.722, 'I': 0.278, 'J': 0.500, 'K': 0.667, 'L': 0.556,
        'M': 0.833, 'N': 0.722, 'O': 0.778, 'P': 0.667, 'Q': 0.778, 'R': 0.722,
        'S': 0.667, 'T': 0.611, 'U': 0.722, 'V': 0.667, 'W': 0.944, 'X': 0.667,
        'Y': 0.667, 'Z': 0.611,
        '[': 0.278, '\\': 0.278, ']': 0.278, '^': 0.469, '_': 0.556, '`': 0.333,
        'a': 0.556, 'b': 0.556, 'c': 0.500, 'd': 0.556, 'e': 0.556, 'f': 0.278,
        'g': 0.556, 'h': 0.556, 'i': 0.222, 'j': 0.222, 'k': 0.500, 'l': 0.222,
        'm': 0.833, 'n': 0.556, 'o': 0.556, 'p': 0.556, 'q': 0.556, 'r': 0.333,
        's': 0.500, 't': 0.278, 'u': 0.556, 'v': 0.500, 'w': 0.722, 'x': 0.500,
        'y': 0.500, 'z': 0.500,
        '{': 0.334, '|': 0.260, '}': 0.334, '~': 0.584,
    }
    _ARIAL_EM_DEFAULT = 0.556  # avg ASCII fallback for missing chars
    _CJK_EM = 1.0              # CJK glyphs are roughly full-width (1 em)

    def _char_em(ch):
        """Real Arial em-width for one character; CJK chars are 1.0 em."""
        if _is_cjk(ch):
            return _CJK_EM
        return _ARIAL_EM.get(ch, _ARIAL_EM_DEFAULT)

    def _text_px_width(text, font_size):
        """Sum of real Arial pixel widths across the whole string."""
        return sum(_char_em(c) for c in text) * font_size

    def _wrap_text(text, max_chars, max_lines_limit, *, font_size=None):
        """Wrap text into lines.

        Two measurement modes:
          • PIXEL MODE  (font_size given AND > 0):
            `max_chars` is treated as a PIXEL budget (max line width).
            Each character's real Arial em-width is accumulated as we
            walk the string; the moment cumulative width would exceed
            the budget we break the line. This is the "wall-bump → break"
            behaviour: the line breaks the instant a character would push
            past the block edge, regardless of which characters they are.
            Wide letters (M, W) are budgeted as 0.83-0.94 em, narrow
            letters (i, l, j) as 0.22 em — so a line of "Wii" budgets
            very differently from a line of "MMM" even though both are
            3 chars. This guarantees no overflow, ever, for any text.

          • CHARACTER MODE  (font_size None — legacy):
            `max_chars` is a character count; CJK = 1.7 each, others = 1.0.
            Same as pre-pixel-mode behaviour, kept for callers that haven't
            been migrated.

        When no natural separator (space, comma, paren, slash, etc.) is
        available near the break point AND we're splitting an English
        word mid-character, a hyphen is appended — so "Recreational"
        becomes "Recreatio-" / "nal" instead of "Recreat…".
        """
        # Mode-specific measurement helpers.
        if font_size is not None and font_size > 0:
            # Pixel mode: budget is max_chars (px), measure with real em widths.
            def _ch_w(ch):
                return _char_em(ch) * font_size
            def _measure(s):
                return _text_px_width(s, font_size)
        else:
            # Character mode: budget is max_chars (chars), CJK ~ 1.7.
            def _ch_w(ch):
                return 1.7 if _is_cjk(ch) else 1.0
            def _measure(s):
                return _visual_len(s)
        budget = max_chars

        if _measure(text) <= budget:
            return [text]

        def _is_ascii_letter(ch):
            return ch.isalpha() and not _is_cjk(ch)

        lines = []
        remaining = text
        # Track where the current iteration's input starts in the ORIGINAL `text`.
        # After the loop, this holds the start position of the LAST kept line,
        # which the fallback truncation uses to compute a proper hyphenated cut
        # for multi-line wrapping. (See the fallback block below for details.)
        last_line_text_start = 0

        while remaining and len(lines) < max_lines_limit:
            last_line_text_start = len(text) - len(remaining)
            if _measure(remaining) <= budget:
                lines.append(remaining)
                remaining = ""
                break

            # Find break point: walk forward measuring real width.
            # In pixel mode: width is the actual rendered px sum.
            # In character mode: width is the legacy CJK-weighted count.
            break_point = 0
            vlen = 0.0
            for ci, ch in enumerate(remaining):
                vlen += _ch_w(ch)
                if vlen > budget:
                    break
                break_point = ci + 1

            if break_point == 0:
                break_point = 1

            # Prefer natural separators
            best_sep = -1
            min_pos = max(1, break_point // 3)
            for si in range(break_point - 1, min_pos - 1, -1):
                ch = remaining[si]
                if ch in (' ', ',', '_', '-', '(', ')', '/', '、', '，'):
                    best_sep = si + 1
                    break
                if si > 0 and (_is_cjk(remaining[si]) != _is_cjk(remaining[si - 1])):
                    best_sep = si
                    break

            if best_sep > 0:
                break_point = best_sep
                # Normal word-boundary break: no hyphen needed.
                line = remaining[:break_point].rstrip()
            else:
                # No natural separator found — we're mid-word. Try to
                # emit a hyphenated break: reserve the last visible slot
                # for the hyphen so the visual line is still ≤ budget.
                # Applies only when BOTH sides of the break are ASCII
                # letters (avoids hyphenating across a digit/punct or CJK).
                # The "≥ 3" lower bound on budget is a sanity check —
                # if the budget is so small even the hyphen wouldn't fit,
                # skip hyphenation. In char mode this is "≥ 3 chars"; in
                # pixel mode it's "≥ 3 em-widths-worth-of-px".
                min_budget_for_hyphen = 3 if font_size is None else 3 * _ARIAL_EM_DEFAULT * font_size
                can_hyphenate = (
                    budget >= min_budget_for_hyphen
                    and break_point >= 2
                    and break_point < len(remaining)
                    and _is_ascii_letter(remaining[break_point - 1])
                    and _is_ascii_letter(remaining[break_point])
                )
                if can_hyphenate:
                    hyph_break = break_point - 1
                    prefix = remaining[:hyph_break].rstrip()
                    if prefix:
                        line = prefix + "-"
                        break_point = hyph_break
                    else:
                        line = remaining[:break_point].rstrip()
                else:
                    line = remaining[:break_point].rstrip()

            # Hard-cap: if the line still exceeds budget (e.g. single
            # character overrun), force-truncate with ellipsis rather than
            # overflow. In practice the hyphenation branch above already
            # keeps us ≤ budget, so this is only a fallback.
            if _measure(line) > budget:
                # Trim from the end until it fits, then add the ellipsis.
                while line and _measure(line + "…") > budget:
                    line = line[:-1]
                line = (line + "…") if line else "…"
            lines.append(line)
            remaining = remaining[break_point:].lstrip()

        # If there's remaining text and we've hit max lines, we MUST truncate.
        # Never silently drop text — always show ellipsis.
        if remaining and len(lines) == max_lines_limit:
            last_line = lines[-1]
            if _measure(last_line) + _measure(remaining) <= budget:
                # All remaining fits appended to the last line — just join them.
                lines[-1] = last_line + remaining
            else:
                # Truncate the last line with an ellipsis. (See long comment
                # in earlier versions for why the "-" → "…" distinction
                # matters here: we use "…" to signal "content was DROPPED",
                # while the in-loop "-" signals "continued on next line".)
                if last_line.endswith("-"):
                    last_line = last_line[:-1]
                # Trim from the end until last_line + "…" fits.
                while last_line and _measure(last_line + "…") > budget:
                    last_line = last_line[:-1]
                lines[-1] = (last_line + "…") if last_line else "…"

        return lines if lines else [text]

    def _hex_to_rgba(hex_color, opacity):
        hex_color = str(hex_color).strip()
        if not hex_color.startswith("#"):
            hex_color = "#" + hex_color
        if len(hex_color) != 7:
            return f"rgba(200,200,200,{opacity})"
        try:
            r = int(hex_color[1:3], 16)
            g = int(hex_color[3:5], 16)
            b = int(hex_color[5:7], 16)
            return f"rgba({r},{g},{b},{opacity})"
        except Exception:
            return f"rgba(200,200,200,{opacity})"

    for sl in slices:
        sid = sl.slice_id
        cats = [EXO_NAME] + order.get(sid, sl.categories_in_order)
        for c in cats:
            bk = BlockKey(sid, c)
            x0, x1, y0, y1 = layout.block_bbox[bk]

            labels_all = block_labels.get(bk, {"Unknown"})
            labels_evt = set(labels_all) & _BLOCK_EVENT_SET
            if not labels_evt:
                labels_evt = {"Unknown"}
            event_text = format_block_event_text(labels_evt)
            primary = _primary_block_label(labels_evt)
            block_name = disp_cat(c)

            elements_line = ""
            if c != EXO_NAME:
                elems = sorted(list(sl.cat_to_elems.get(c, set())))
                elements_line = f"<br>Elements ({len(elems)}): " + ", ".join(elems)

            # Hover tooltip text construction.
            # For Exo blocks: just the bold name, nothing else. Exo represents
            # "outside-the-grouping-system" elements with no event semantics
            # of their own — adding the event text would be confusing/noisy
            # since it doesn't describe an actual category event. Per user
            # request, Exo gets a minimal "Exo"-only tooltip.
            # For all other blocks: bold name + event text + elements list.
            if c == EXO_NAME:
                hover_html_block = f"<b>{block_name}</b>"
            else:
                hover_html_block = f"<b>{block_name}</b><br>{event_text}{elements_line}"

            # Get per-type style settings (fallback to defaults)
            fill_colors_map = getattr(cfg, "block_fill_colors_by_type", {}) or {}
            fill_opacities_map = getattr(cfg, "block_fill_opacities_by_type", {}) or {}
            border_colors_map = getattr(cfg, "block_border_colors_by_type", {}) or {}
            border_opacities_map = getattr(cfg, "block_border_opacities_by_type", {}) or {}
            border_widths_map = getattr(cfg, "block_border_widths_by_type", {}) or {}
            line_styles_map = getattr(cfg, "block_line_styles_by_type", {}) or {}
            text_fonts_map = getattr(cfg, "block_text_fonts_by_type", {}) or {}
            text_sizes_map = getattr(cfg, "block_text_sizes_by_type", {}) or {}
            text_colors_map = getattr(cfg, "block_text_colors_by_type", {}) or {}
            text_aligns_map = getattr(cfg, "block_text_aligns_by_type", {}) or {}
            text_rotations_map = getattr(cfg, "block_text_rotations_by_type", {}) or {}
            line_spacings_map = getattr(cfg, "block_line_spacings_by_type", {}) or {}
            border_radii_map = getattr(cfg, "block_border_radii_by_type", {}) or {}

            # Get styles for primary event type
            fill_color = fill_colors_map.get(str(primary),
                                             cfg.block_colors.get(primary, cfg.block_colors.get("Unknown", "#CCCCCC")))
            fill_opacity = float(fill_opacities_map.get(str(primary), 1.0))
            border_color = border_colors_map.get(str(primary), "#222222")
            border_opacity = float(border_opacities_map.get(str(primary), 1.0))
            border_width = float(border_widths_map.get(str(primary), cfg.block_border_width))
            line_style = line_styles_map.get(str(primary), "solid")
            text_font = text_fonts_map.get(str(primary), "Arial")
            text_size = int(text_sizes_map.get(str(primary), cfg.block_text_size))
            text_color = text_colors_map.get(str(primary),
                                             cfg.block_text_colors.get(primary, cfg.block_text_color_default))
            text_align = text_aligns_map.get(str(primary), "center")
            text_rotation = int(text_rotations_map.get(str(primary), 90))
            line_spacing = float(line_spacings_map.get(str(primary), 0))
            border_radii = border_radii_map.get(str(primary), [0, 0, 0, 0])

            # Ensure border_radii is a valid list with no None values
            if not border_radii or len(border_radii) != 4:
                border_radii = [0, 0, 0, 0]
            border_radii = [0 if r is None else r for r in border_radii]

            # Generate rounded rectangle coordinates if radii are set
            # radii order: [top-left, top-right, bottom-right, bottom-left]
            # Convert to data units (radii are in pixels, need to scale)
            # Use a simple scaling factor based on block dimensions
            radii_scaled = border_radii
            if any(r > 0 for r in border_radii):
                # Scale radii from pixel-like values to data units
                # Assume radii values are roughly in the same scale as block dimensions
                block_w = x1 - x0
                block_h = y1 - y0
                scale_factor = min(block_w, block_h) / 50.0  # 50 is max radius in UI
                radii_scaled = [(r if r is not None else 0) * scale_factor for r in border_radii]

            xs, ys = rounded_rect_polygon_coords(x0, y0, x1, y1, radii_scaled)

            # Convert line style to dash pattern
            dash_pattern = None
            if line_style == "dash":
                dash_pattern = "dash"
            elif line_style == "dot":
                dash_pattern = "dot"
            elif line_style == "dashdot":
                dash_pattern = "dashdot"
            # "solid" -> None (default)

            fill_color_rgba = _hex_to_rgba(fill_color, fill_opacity)
            border_color_rgba = _hex_to_rgba(border_color, border_opacity)

            # A1) 可见矩形：负责画出来 + 负责 hover（hover 任意位置都显示同一个信息框）
            line_dict = dict(color=border_color_rgba, width=border_width)
            if dash_pattern:
                line_dict["dash"] = dash_pattern

            fig.add_trace(go.Scatter(
                x=xs, y=ys,
                mode="lines",
                name="",
                line=line_dict,
                fill="toself",
                fillcolor=fill_color_rgba,
                opacity=1.0,
                hoveron="fills",
                hoverinfo="text",
                text=hover_html_block,
                hovertemplate="%{text}<extra></extra>",
                showlegend=False,
            ))
            curve_to_id.append(block_id(bk))

            # --- REQUEST #3: Exo text centered (same as other blocks) ---
            cx, cy = layout.block_xy[bk]
            title = "Exo" if c == EXO_NAME else c

            # --- Text wrapping: physically-correct, viewport-adaptive ---
            #
            # Plotly annotation font size is in fixed screen-pixels, but block
            # dimensions are in data-units.  With scaleanchor="x",scaleratio=1
            # both axes share one uniform px_per_unit:
            #
            #   px_per_unit = min(plot_W / x_span,  plot_H / y_span)
            #
            # We estimate plot_W / plot_H conservatively, then compute:
            #   max_chars = char_dim_data * px_per_unit / (font_size * char_w_ratio)
            #   max_lines = line_dim_data * px_per_unit / (font_size * line_spacing)
            #
            # Both constraints are enforced so text never overflows in either direction.
            # --- Apply text padding (L, R, T, B) from cfg ---
            _tp_l = getattr(cfg, 'block_text_pad_l', 0.0)
            _tp_r = getattr(cfg, 'block_text_pad_r', 0.0)
            _tp_t = getattr(cfg, 'block_text_pad_t', 0.0)
            _tp_b = getattr(cfg, 'block_text_pad_b', 0.0)

            # Padded bounds for text area
            _px0 = x0 + _tp_l
            _px1 = x1 - _tp_r
            _py0 = y0 + _tp_b
            _py1 = y1 - _tp_t

            # Apply fractional safety inset on TOP of the per-side pad_* values.
            # IMPORTANT — asymmetric by design:
            #   • LINES-stacking direction gets the inset. Line-height
            #     rendering is the leaky axis: Plotly's declared 1.2×
            #     line-height plus ascender/descender leading is the
            #     only safety we have, and there's no "auto-shrink"
            #     recovery if a line pokes above/below the block.
            #   • CHARS direction gets NO inset. If text would exceed
            #     the available char-direction space, _wrap_text breaks
            #     words with a hyphen (see _wrap_text), so there is no
            #     overflow risk in this axis — any inset here is pure
            #     wasted space, which is exactly what the user flagged.
            # text_rotation is already resolved above (line ~3604).
            _safety_inset = float(getattr(cfg, 'block_text_safety_inset_frac', 0.10) or 0.0)
            if _safety_inset > 0.0:
                _raw_w = _px1 - _px0
                _raw_h = _py1 - _py0
                if abs(text_rotation) > 45:
                    # Vertical text: chars run along block_h, lines stack along block_w.
                    # Apply inset ONLY on x (width = lines direction).
                    _dx = _raw_w * _safety_inset * 0.5
                    _px0 += _dx
                    _px1 -= _dx
                    # y axis (chars direction) unchanged — full block_h available.
                else:
                    # Horizontal text: chars run along block_w, lines stack along block_h.
                    # Apply inset ONLY on y (height = lines direction).
                    _dy = _raw_h * _safety_inset * 0.5
                    _py0 += _dy
                    _py1 -= _dy
                    # x axis (chars direction) unchanged.

            # Clamp so padded area doesn't invert
            if _px1 <= _px0:
                _px0 = (x0 + x1) / 2.0 - 0.001
                _px1 = (x0 + x1) / 2.0 + 0.001
            if _py1 <= _py0:
                _py0 = (y0 + y1) / 2.0 - 0.001
                _py1 = (y0 + y1) / 2.0 + 0.001

            block_w = _px1 - _px0
            block_h = _py1 - _py0
            _pcx = (_px0 + _px1) / 2.0
            _pcy = (_py0 + _py1) / 2.0

            _EST_W = _CANVAS_W  # actual plot-area width (px), set from graph_size
            _EST_H = _CANVAS_H  # actual plot-area height (px), set from graph_size
            # _CHAR_W: only used as a quick sanity estimate for max_lines
            # (height-axis safety). The CHARS-axis (line width) decision is
            # now made by _wrap_text in pixel mode using real Arial em
            # widths, so this constant no longer drives line-break
            # accuracy. Kept conservatively at 0.55 (Arial mixed-case avg)
            # purely to compute max_lines without further side effects.
            _CHAR_W = 0.55
            # line_spacing is absolute px offset (0 = Plotly default, +N = wider, -N = tighter)
            _PLOTLY_LH = 1.2  # Plotly's native line-height ratio (back to v44 value)
            _base_lh_px = text_size * _PLOTLY_LH
            _eff_lh_px = max(text_size * 0.3, _base_lh_px + float(line_spacing))
            _LINE_H = _eff_lh_px / max(text_size, 1)  # convert to ratio for max_lines calc
            # Use the EFFECTIVE ppu (real canvas dimensions × contain-fit zoom),
            # matching v44. Earlier in v48 we capped this to a "reference" canvas
            # (_TEXT_FIT_REF_W × _TEXT_FIT_REF_H) for cross-browser stability,
            # but that was also losing 10-15% of available text width on wide
            # screens — undermining the v44 visual density that the user
            # confirmed was the desired baseline. Cross-browser stability is
            # now provided primarily by the 25px quantization on graph_size
            # (see clientside callbacks) which tolerates ±12.5px of measurement
            # jitter without producing different wrap decisions.
            _ppu_x = _eff_ppu_x   # full effective px/unit at initial display
            _ppu_y = _eff_ppu_y

            if abs(text_rotation) > 45:
                # Vertical text (rotation≈90°): characters run along the block HEIGHT (y-axis),
                # wrapped lines stack along the block WIDTH (x-axis).
                char_dim_px = block_h * _ppu_y   # height in px → chars per line
                line_dim_px = block_w * _ppu_x   # width  in px → number of lines
            else:
                # Horizontal text: characters along WIDTH, lines stack along HEIGHT.
                char_dim_px = block_w * _ppu_x
                line_dim_px = block_h * _ppu_y

            # max_lines: a SAFETY CAP for the lines-stacking direction. The
            # `- 1` keeps a 1-line buffer to absorb ±1-line measurement
            # jitter (line-height rounding, plotly leading, etc.). Lines
            # axis still relies on integer-character estimates because
            # there's no equivalent of pixel-accurate width measurement
            # for VERTICAL stacking — line height is dominated by plotly's
            # internal leading which we can't introspect from Python.
            #
            # _LINE_H is the per-line vertical extent ratio used in this
            # safety check. Plotly's declared 1.2× lineheight is a lower
            # bound — actual rendering adds ascender/descender leading
            # that pushes the effective per-line vertical step closer
            # to ~1.5×. Using 1.5 here makes max_lines slightly more
            # conservative, preventing the most-conspicuous failure mode
            # where wrap thinks 6 lines will fit but 5 already touch
            # the block top/bottom edges.
            _LINES_SAFETY = 1.5
            max_lines = max(1, int(line_dim_px / max(text_size * _LINES_SAFETY, 1)) - 1)

            # CHARS axis: pixel-accurate wrap.
            # Pass char_dim_px directly as the budget and tell _wrap_text
            # the font size so it can use the real Arial em-width table.
            # The 0.92 safety factor on the budget gives an ~8% margin
            # between the wrap decision and the actual block edge — covers
            # measurement noise from the viewport-derived _ppu and from
            # plotly's own glyph-positioning rounding. Without this margin,
            # individual blocks where the residual error happens to push
            # the same direction as the wrap decision would render text
            # one or two pixels past the visible edge — visible as
            # "compressed" or "overflowing" labels in the screenshot.
            _CHARS_SAFETY = 0.92
            wrap_budget_px = char_dim_px * _CHARS_SAFETY
            wrapped_lines = _wrap_text(
                title,
                wrap_budget_px,         # budget: pixels of available width (with margin)
                max_lines,              # vertical safety cap
                font_size=text_size,    # enables pixel-accurate measurement
            )

            # Legacy variable for any downstream reference. Equals the
            # "average char-count per line" estimate, only meaningful if
            # something else still expects it; the wrap itself ignores it.
            max_chars_per_line = max(1, int(char_dim_px / max(text_size * _CHAR_W, 1)))

            # Build display text with line-spacing-aware separators
            # Each extra <br> adds ~font_size * 1.2 px of vertical space
            _blk_extra_brs = ""
            if line_spacing > 0 and _base_lh_px > 0:
                _blk_n_extra = max(0, round(float(line_spacing) / _base_lh_px))
                if _blk_n_extra > 0:
                    _blk_extra_brs = "<br>" * _blk_n_extra
            title_display = ("<br>" + _blk_extra_brs).join(wrapped_lines)

            # Determine text position based on alignment — use padded bounds
            if text_align == "left":
                text_x = _px0
                xanchor = "left"
            elif text_align == "right":
                text_x = _px1
                xanchor = "right"
            else:  # center (default)
                text_x = _pcx
                xanchor = "center"

            # Use annotation for text (supports textangle for rotation)
            # Track the annotation index for visibility control
            current_anno_count = len(fig.layout.annotations) if fig.layout.annotations else 0
            fig.add_annotation(
                x=text_x,
                y=_pcy,
                text=title_display,
                showarrow=False,
                xanchor=xanchor,
                yanchor="middle",
                font=dict(size=text_size, color=text_color, family=text_font),
                textangle=-text_rotation,  # Plotly uses opposite sign convention
                bgcolor="rgba(0,0,0,0)",
                bordercolor="rgba(0,0,0,0)",
                borderpad=0,
            )
            # Track block text annotation index for visibility control (e.g., hiding Exo text)
            block_text_anno_indices[block_id(bk)] = current_anno_count

            block_items.append({
                "oid": block_id(bk),
                "xs": xs,
                "ys": ys,
                "hover": hover_html_block,
                "primary": str(primary),
            })

    # Create one clickable "show Exo" button per slice (DEFAULT HIDDEN).
    # It will be shown only when that slice's Exo is hidden.
    for sl in slices:
        sid = sl.slice_id
        exo_bk = BlockKey(sid, EXO_NAME)
        x0, x1, y0, y1 = layout.block_bbox[exo_bk]
        cx, cy = layout.block_xy[exo_bk]
        btn_x = x0 - cfg.exo_toggle_gap
        btn_y = cy

        fig.add_trace(go.Scatter(
            x=[btn_x], y=[btn_y],
            mode="markers+text",
            name=" ",
            marker=dict(
                size=cfg.exo_toggle_marker_size,
                symbol="square",
                color=cfg.exo_toggle_marker_fill,
                line=dict(width=1, color=cfg.exo_toggle_marker_line)
            ),
            text=["E"],
            textposition="middle center",
            textfont=dict(size=cfg.exo_toggle_text_size, color=cfg.exo_toggle_text_color),
            hovertemplate="Show <b>Exo</b><extra></extra>",
            showlegend=False,
            opacity=0.98,
            visible=False,
        ))
        curve_to_id.append(f"BTNEXO::{sid}")
    # --- Slice labels (text only, positioned to the right of outermost enclosure) ---
    slice_anno_indices: Dict[str, int] = {}
    slice_label_base_x: Dict[str, float] = {}
    slice_label_base_y: Dict[str, float] = {}
    # Slice label Distance = gap from label left edge to outermost enclosure right edge.
    # Default distance stored in slice_style; base_x is the enclosure right edge.
    _SLICE_LABEL_DEFAULT_DIST = 2.0  # default distance in data units
    slice_label_anchor_x = enclosure_global_x1 + _SLICE_LABEL_DEFAULT_DIST
    for sl in slices:
        sid = sl.slice_id
        exo_bk = BlockKey(sid, EXO_NAME)
        _, cy = layout.block_xy[exo_bk]
        x_label = float(slice_label_anchor_x)
        y_label = float(cy)

        # base_x stores the enclosure right edge (reference point for distance control)
        slice_label_base_x[str(sid)] = float(enclosure_global_x1)
        slice_label_base_y[str(sid)] = float(y_label)

        cur_n = len(fig.layout.annotations) if fig.layout.annotations else 0
        fig.add_annotation(
            x=x_label,
            y=y_label,
            text=str(sid),
            showarrow=False,
            xanchor="left",
            yanchor="middle",
            font=dict(size=23, color="#000000", family="Arial"),
            textangle=0,
            bgcolor="rgba(0,0,0,0)",
            bordercolor="rgba(0,0,0,0)",
            borderpad=0,
            visible=True,
            templateitemname=f"SLICE_LABEL::{sid}",
        )
        slice_anno_indices[str(sid)] = int(cur_n)

    # --- Collapse state info (no dots drawn on graph; checkboxes are in the panel) ---
    is_collapsed = (collapse_state or {}).get("collapsed", False)
    orig_anchor_x = (collapse_state or {}).get("anchor_x")
    orig_anchor_y = (collapse_state or {}).get("anchor_y")
    collapse_line_hidden = (collapse_state or {}).get("line_hidden", False)

    # --- Collapse line (shown when collapsed, between anchor slices) ---
    # Drawn as a Plotly shape (not a trace) so it appears in downloaded images.
    # Visibility controlled by collapse_line_hidden flag.
    collapse_line_shape_idx = None
    if is_collapsed and len(slices) >= 2 and orig_anchor_x is not None:
        anchor_x_view_idx = orig_anchor_x
        anchor_y_view_idx = orig_anchor_x + 1

        if anchor_x_view_idx < len(slices) and anchor_y_view_idx < len(slices):
            anchor_x_sid = slices[anchor_x_view_idx].slice_id
            anchor_y_sid = slices[anchor_y_view_idx].slice_id

            exo_bk_x = BlockKey(anchor_x_sid, EXO_NAME)
            exo_bk_y = BlockKey(anchor_y_sid, EXO_NAME)

            if exo_bk_x in layout.block_xy and exo_bk_y in layout.block_xy:
                _, y_x = layout.block_xy[exo_bk_x]
                _, y_y = layout.block_xy[exo_bk_y]

                collapse_y = (y_x + y_y) / 2.0
                collapse_line_x0 = x_range[0] + float(cfg.viewport_pad_x)
                # Right end = left edge of slice label (exact anchor position)
                collapse_line_x1 = float(slice_label_anchor_x)

                collapse_line_shape_idx = None  # Will be set after adding to figure
                _collapse_line_shape = dict(
                    type="line",
                    x0=collapse_line_x0, y0=collapse_y,
                    x1=collapse_line_x1, y1=collapse_y,
                    line=dict(color="rgba(100, 100, 100, 0.6)", width=3, dash="dash"),
                    layer="above",
                    visible=(not collapse_line_hidden),
                )
                # Add shape to figure (shapes already applied, so extend via update_layout)
                _existing_shapes = list(fig.layout.shapes or [])
                collapse_line_shape_idx = len(_existing_shapes)
                fig.update_layout(shapes=_existing_shapes + [_collapse_line_shape])

    # --- Block hover/click hitboxes (draw AFTER bands so blocks remain hoverable) ---
    for it in block_items:
        fig.add_trace(go.Scatter(
            x=it["xs"], y=it["ys"],
            mode="lines",
            name="",
            line=dict(color="rgba(0,0,0,0)", width=0),
            fill="toself",
            fillcolor="rgba(0,0,0,0)",
            opacity=0.001,
            hoveron="fills",
            hoverinfo="text",
            text=it["hover"],
            hovertemplate="%{text}<extra></extra>",
            showlegend=False,
        ))
        curve_to_id.append(it["oid"])

    fig.update_xaxes(range=x_range)
    fig.update_yaxes(range=y_range, scaleanchor="x", scaleratio=1)

    if not cfg.show_axes:
        fig.update_xaxes(visible=False)
        fig.update_yaxes(visible=False)
    fig.update_xaxes(showspikes=False)
    fig.update_yaxes(showspikes=False)

    levels_out: List[Dict[str, str]] = []
    seen_keys: Set[str] = set()
    keys_present = set(label_level_keys_all)

    if slices:
        max_lv = max((len(sl.agg_cols) for sl in slices), default=0)

        for idx_lv in range(max_lv):
            key_lv = "group" if idx_lv == 0 else f"supergroup{idx_lv}"
            if key_lv not in keys_present:
                continue
            if key_lv in seen_keys:
                continue

            col_lv = None
            for sl in slices:
                if idx_lv < len(sl.agg_cols):
                    col_lv = sl.agg_cols[idx_lv]
                    break
            if col_lv is None:
                continue

            seen_keys.add(key_lv)
            label_lv = "Group level" if idx_lv == 0 else f"Supergroup{idx_lv} level"
            levels_out.append({
                "key": key_lv,
                "col": str(col_lv),
                "label": f"{label_lv} ({col_lv})"
            })

    # ── Truncated label tooltip map (consumed by render_with_highlight) ────
    # Keys: annotation index (str) → original full text.
    # render_with_highlight rebuilds the FINAL text→full map AFTER
    # apply_enclosure_label_visibility rewrites/wraps the annotations, so the
    # JS proximity-check map always matches the actual SVG textContent.
    _truncated_label_by_anno_idx: dict = {}
    for spec, anno_idx in zip(label_specs_all, label_anno_indices):
        orig = spec.get("orig_text") or spec.get("full_text") or spec["text"]
        _truncated_label_by_anno_idx[str(anno_idx)] = orig

    meta = {
        "curve_to_id": curve_to_id,
        "initial_x_range": list(x_range),  # baseline for zoom-scale calculation
        "initial_y_range": list(y_range),  # preserved across collapse rebuilds
        "block_to_bands": {k: sorted(list(v)) for k, v in block_to_bands.items()},
        "band_to_blocks": {k: sorted(list(v)) for k, v in band_to_blocks.items()},
        "slices": [sl.slice_id for sl in slices],
        "slice_label_anno_indices": dict(slice_anno_indices),
        "slice_label_base_x": dict(slice_label_base_x),
        "slice_label_base_y": dict(slice_label_base_y),
        "enclosure_global_x0": float(enclosure_global_x0),
        "enclosure_global_x1": float(enclosure_global_x1),
        "y_step": float(layout.y_step),
        "collapse_line_shape_idx": collapse_line_shape_idx,
        "band_shape_indices": dict(band_shape_indices),
        "band_primary_by_oid": {str(it.get("bid")): str(it.get("primary")) for it in band_items},
        "block_cx": {f"{bk.slice_id}::{bk.cat_name}": float((bb[0] + bb[1]) / 2.0) for bk, bb in
                     layout.block_bbox.items()},
        "block_text_anno_indices": dict(block_text_anno_indices),
        "enclosure_levels": levels_out,
        "enclosure_label_level_keys": list(label_level_keys_all),
        "enclosure_label_mask_shape_indices": list(mask_shape_indices),
        "enclosure_label_border_shape_indices": list(border_shape_indices),
        "enclosure_label_anno_indices": list(label_anno_indices),
        "enclosure_label_depths": list(label_depths),
        "truncated_label_by_anno_idx": _truncated_label_by_anno_idx,
    }

    return fig, meta


def apply_highlight(fig_json: dict, meta: dict, active_id: Optional[str], cfg: StyleConfig) -> dict:
    """
    Requested 3-state highlight:
      - CORE (selected): stronger (darker color)
      - CONNECTED (reachable in Block<->Band graph, stop expanding at Exo but include Exo itself): unchanged
      - UNCONNECTED: dimmed by opacity
    The state persists because we always render from base_fig + selected_id.
    """
    if not active_id:
        return fig_json

    curve_to_id = meta.get("curve_to_id", [])
    block_to_bands = meta.get("block_to_bands", {})
    band_to_blocks = meta.get("band_to_blocks", {})

    def _safe_float(v: Any, default: float) -> float:
        try:
            return float(v)
        except Exception:
            return float(default)

    def is_exo_block(oid: str) -> bool:
        parts = oid.split("::")
        return len(parts) >= 3 and parts[0] == "BLOCK" and parts[2] == EXO_NAME

    # --- "No-backtracking" highlight (bidirectional) ---
    # Rule:
    #   - Forward-only expansion: follow BAND direction src_slice -> dst_slice (time increases)
    #   - Backward-only expansion: follow reverse direction (time decreases)
    #   - Do NOT mix directions (i.e., never go backward then forward, or forward then backward)

    slice_order = meta.get("slices", [])
    slice_to_idx = {sid: i for i, sid in enumerate(slice_order)}

    def parse_block(oid: str) -> Tuple[str, str]:
        # "BLOCK::slice::cat"
        parts = oid.split("::")
        if len(parts) < 3:
            return ("", "")
        return (parts[1], parts[2])

    def parse_band(oid: str) -> Tuple[str, str, str, str]:
        # "BAND::src_slice::src_cat::dst_slice::dst_cat"
        parts = oid.split("::")
        if len(parts) < 5:
            return ("", "", "", "")
        return (parts[1], parts[2], parts[3], parts[4])

    def make_block(sid: str, cat: str) -> str:
        return f"BLOCK::{sid}::{cat}"

    def outgoing_bands_of_block(block_oid: str) -> List[str]:
        sid, cat = parse_block(block_oid)
        out: List[str] = []
        for band_oid in block_to_bands.get(block_oid, []):
            if not band_oid.startswith("BAND::"):
                continue
            s_sid, s_cat, d_sid, d_cat = parse_band(band_oid)
            if (s_sid == sid) and (s_cat == cat):
                si = slice_to_idx.get(s_sid, None)
                di = slice_to_idx.get(d_sid, None)
                if (si is not None) and (di is not None) and (di > si):
                    out.append(band_oid)
        return out

    def incoming_bands_of_block(block_oid: str) -> List[str]:
        sid, cat = parse_block(block_oid)
        inc: List[str] = []
        for band_oid in block_to_bands.get(block_oid, []):
            if not band_oid.startswith("BAND::"):
                continue
            s_sid, s_cat, d_sid, d_cat = parse_band(band_oid)
            if (d_sid == sid) and (d_cat == cat):
                si = slice_to_idx.get(s_sid, None)
                di = slice_to_idx.get(d_sid, None)
                if (si is not None) and (di is not None) and (di > si):
                    inc.append(band_oid)
        return inc

    blocks_on: Set[str] = set()
    bands_on: Set[str] = set()

    forward_starts: List[str] = []
    backward_starts: List[str] = []

    if active_id.startswith("BLOCK::"):
        blocks_on.add(active_id)
        if not is_exo_block(active_id):
            forward_starts.append(active_id)
            backward_starts.append(active_id)

    elif active_id.startswith("BAND::"):
        bands_on.add(active_id)
        s_sid, s_cat, d_sid, d_cat = parse_band(active_id)

        src_block = make_block(s_sid, s_cat)
        dst_block = make_block(d_sid, d_cat)

        blocks_on.add(src_block)
        blocks_on.add(dst_block)

        if not is_exo_block(dst_block):
            forward_starts.append(dst_block)
        if not is_exo_block(src_block):
            backward_starts.append(src_block)

    else:
        return fig_json

    # Forward-only walk
    qf: List[str] = list(forward_starts)
    vf: Set[str] = set()
    while qf:
        b = qf.pop()
        if b in vf:
            continue
        vf.add(b)

        if is_exo_block(b):
            continue

        for band_oid in outgoing_bands_of_block(b):
            bands_on.add(band_oid)
            s_sid, s_cat, d_sid, d_cat = parse_band(band_oid)
            nb = make_block(d_sid, d_cat)
            blocks_on.add(nb)
            if (not is_exo_block(nb)) and (nb not in vf):
                qf.append(nb)

    # Backward-only walk
    qb: List[str] = list(backward_starts)
    vb: Set[str] = set()
    while qb:
        b = qb.pop()
        if b in vb:
            continue
        vb.add(b)

        if is_exo_block(b):
            continue

        for band_oid in incoming_bands_of_block(b):
            bands_on.add(band_oid)
            s_sid, s_cat, d_sid, d_cat = parse_band(band_oid)
            nb = make_block(s_sid, s_cat)
            blocks_on.add(nb)
            if (not is_exo_block(nb)) and (nb not in vb):
                qb.append(nb)

    fig = go.Figure(fig_json)

    def is_block_hitbox_trace(tr) -> bool:
        """
        Your block hitboxes are invisible polygons (opacity ~ 0.001, fillcolor rgba(0,0,0,0), line width 0).
        We must NEVER let them become visible by opacity changes.
        """
        try:
            if getattr(tr, "fillcolor", None) != "rgba(0,0,0,0)":
                return False
            op = float(getattr(tr, "opacity", 1.0))
            if op > 0.01:
                return False
            ln = getattr(tr, "line", None)
            if ln is None:
                return False
            if getattr(ln, "width", None) != 0:
                return False
            return True
        except Exception:
            return False

    for idx, tr in enumerate(fig.data):
        if idx >= len(curve_to_id):
            continue

        oid = curve_to_id[idx]

        # Keep background & exo buttons always available
        if oid == "BG::CLICK" or oid.startswith("BTNEXO::"):
            tr.opacity = 1.0
            tr.visible = True
            continue

        is_text = oid.endswith("::TEXT")
        base_oid = oid.replace("::TEXT", "") if is_text else oid

        # Never touch hitbox traces (avoid making them visible)
        if is_block_hitbox_trace(tr):
            continue

        # --- BLOCKS ---
        if base_oid.startswith("BLOCK::"):
            if base_oid == active_id:
                # CORE: keep original colors; strengthen via emboss + inner glow overlays (added later)
                tr.opacity = 1.0

                if (not is_text) and hasattr(tr, "line") and tr.line:
                    tr.line.width = cfg.core_block_border_width

            elif base_oid in blocks_on:
                # CONNECTED: unchanged
                continue

            else:
                # UNCONNECTED: dim
                tr.opacity = cfg.dim_opacity

            continue

        # --- BANDS ---
        if base_oid.startswith("BAND::"):
            if base_oid == active_id:
                # CORE: keep original colors; strengthen via emboss + inner glow overlays (added later)
                tr.opacity = 1.0

                if hasattr(tr, "line") and tr.line:
                    tr.line.width = cfg.core_band_line_width

            elif base_oid in bands_on:
                # CONNECTED: unchanged
                continue

            else:
                # UNCONNECTED: dim
                tr.opacity = cfg.dim_opacity

            continue

        # --- NEW: highlight visible band SHAPES as well (since band polygons are shapes, not traces) ---
    try:
        band_shape_indices = meta.get("band_shape_indices", {}) or {}
        if band_shape_indices and hasattr(fig.layout, "shapes") and fig.layout.shapes:
            for boid, si in band_shape_indices.items():
                try:
                    si_int = int(si)
                except Exception:
                    continue
                if si_int < 0 or si_int >= len(fig.layout.shapes):
                    continue
                if not str(boid).startswith("BAND::"):
                    continue

                shp = fig.layout.shapes[si_int]

                if str(boid) == active_id:
                    # CORE: opacity 1.0; side-border overlays added later.
                    shp.opacity = 1.0
                    try:
                        ln = getattr(shp, "line", None)
                        if ln is not None:
                            ln.width = cfg.core_band_line_width
                    except Exception:
                        pass

                elif str(boid) in bands_on:
                    # CONNECTED: unchanged
                    continue
                else:
                    # UNCONNECTED: dim
                    shp.opacity = float(cfg.dim_opacity)
    except Exception:
        pass

    # --- NEW: dim enclosures (group + supergroup shapes) as well ---
    # Rule: if an enclosure contains NO connected blocks (blocks_on), dim it.
    try:
        block_bbox_by_oid: Dict[str, Tuple[float, float, float, float]] = {}

        # Collect visible block-rectangle bboxes from traces
        for idx2, tr2 in enumerate(fig.data):
            if idx2 >= len(curve_to_id):
                continue

            oid2 = curve_to_id[idx2]
            if oid2.endswith("::TEXT"):
                continue
            if not oid2.startswith("BLOCK::"):
                continue
            if is_block_hitbox_trace(tr2):
                continue

            fc2 = getattr(tr2, "fillcolor", None)
            if fc2 is None or fc2 == "rgba(0,0,0,0)":
                continue

            try:
                xs2 = list(tr2.x) if tr2.x is not None else []
                ys2 = list(tr2.y) if tr2.y is not None else []
            except Exception:
                continue

            if not xs2 or not ys2:
                continue

            block_bbox_by_oid[oid2] = (min(xs2), max(xs2), min(ys2), max(ys2))

        # Dim shapes whose enclosed blocks are all unconnected
        if hasattr(fig.layout, "shapes") and fig.layout.shapes:
            eps = 1e-6
            for si in range(len(fig.layout.shapes)):
                shp = fig.layout.shapes[si]
                try:
                    if getattr(shp, "type", "") not in ("rect", "path"):
                        continue
                    # Skip white label masks (line width 0, fillcolor white)
                    try:
                        fc = str(getattr(shp, "fillcolor", ""))
                        lw = _safe_float(getattr(getattr(shp, "line", None), "width", 1.0), 1.0)
                        if (fc.strip().upper() == "#FFFFFF") and (lw <= 0.0):
                            continue
                    except Exception:
                        pass

                    rx0 = getattr(shp, "x0", None)
                    rx1 = getattr(shp, "x1", None)
                    ry0 = getattr(shp, "y0", None)
                    ry1 = getattr(shp, "y1", None)
                    if (rx0 is None) or (rx1 is None) or (ry0 is None) or (ry1 is None):
                        continue
                    x0 = _safe_float(rx0, 0.0)
                    x1 = _safe_float(rx1, 0.0)
                    y0 = _safe_float(ry0, 0.0)
                    y1 = _safe_float(ry1, 0.0)

                    rx0, rx1 = (x0, x1) if x0 <= x1 else (x1, x0)
                    ry0, ry1 = (y0, y1) if y0 <= y1 else (y1, y0)

                    has_connected_block = False
                    for boid, bb in block_bbox_by_oid.items():
                        bx0, bx1, by0, by1 = bb
                        inside = (
                                (bx0 >= rx0 - eps) and (bx1 <= rx1 + eps) and
                                (by0 >= ry0 - eps) and (by1 <= ry1 + eps)
                        )
                        if not inside:
                            continue
                        if boid in blocks_on:
                            has_connected_block = True
                            break

                    if not has_connected_block:
                        fig.layout.shapes[si].opacity = float(cfg.dim_opacity)


                except Exception:
                    continue
    except Exception:
        pass

    # --- NEW: dim block text annotations as well ---
    # When a block is dimmed, its text annotation should also be dimmed
    try:
        block_text_anno_indices = meta.get("block_text_anno_indices", {}) or {}
        if block_text_anno_indices and hasattr(fig.layout, "annotations") and fig.layout.annotations:
            for block_oid, aidx in block_text_anno_indices.items():
                try:
                    aidx_int = int(aidx)
                except Exception:
                    continue
                if aidx_int < 0 or aidx_int >= len(fig.layout.annotations):
                    continue

                # Check if this block is the active selection, connected, or unconnected
                if str(block_oid) == active_id:
                    # CORE: full opacity
                    fig.layout.annotations[aidx_int].opacity = 1.0
                elif str(block_oid) in blocks_on:
                    # CONNECTED: unchanged (keep default opacity)
                    pass
                else:
                    # UNCONNECTED: dim
                    fig.layout.annotations[aidx_int].opacity = float(cfg.dim_opacity)
    except Exception:
        pass

    # --- NEW: dim enclosure (group/supergroup) label annotations as well ---
    # When an enclosure shape is dimmed, its label annotation should also be dimmed
    try:
        border_indices = list(meta.get("enclosure_label_border_shape_indices", []))
        anno_indices = list(meta.get("enclosure_label_anno_indices", []))

        if border_indices and anno_indices and hasattr(fig.layout, "shapes") and hasattr(fig.layout, "annotations"):
            n = min(len(border_indices), len(anno_indices))
            for i in range(n):
                try:
                    bidx = int(border_indices[i])
                    aidx = int(anno_indices[i])
                except Exception:
                    continue

                if bidx < 0 or bidx >= len(fig.layout.shapes):
                    continue
                if aidx < 0 or aidx >= len(fig.layout.annotations):
                    continue

                # Check if this shape was dimmed
                shape_opacity = getattr(fig.layout.shapes[bidx], "opacity", 1.0)
                try:
                    shape_opacity = float(shape_opacity) if shape_opacity is not None else 1.0
                except Exception:
                    shape_opacity = 1.0

                # If shape is dimmed, also dim the label annotation
                if shape_opacity < 0.5:  # Consider dimmed if opacity < 0.5
                    fig.layout.annotations[aidx].opacity = float(cfg.dim_opacity)
    except Exception:
        pass

    # --- NEW: CORE highlight effect (emboss + inner glow) ---
    # This is applied only to the CORE (clicked) object to make it stand out more strongly.
    try:
        def _add_core_rect_fx(x0: float, x1: float, y0: float, y1: float) -> None:
            w = max(1e-9, float(x1 - x0))
            h = max(1e-9, float(y1 - y0))

            # Stronger inner glow: more layers, higher opacity, slightly thicker strokes.
            inset_step = min(w, h) * 0.065
            glow_line_w = max(3.0, float(cfg.core_block_border_width) * 2.4)

            # Use several inset rectangles to mimic an "inner" glow.
            for k in range(1, 6):
                inset = (k * inset_step) / 5.0
                xi0 = x0 + inset
                xi1 = x1 - inset
                yi0 = y0 + inset
                yi1 = y1 - inset
                if (xi1 <= xi0) or (yi1 <= yi0):
                    break

                # Brighter near the edge, softer toward the center.
                alpha = 0.85 / float(k)
                fig.add_shape(
                    type="rect",
                    x0=xi0,
                    x1=xi1,
                    y0=yi0,
                    y1=yi1,
                    line=dict(color=f"rgba(255,255,255,{alpha})", width=glow_line_w),
                    fillcolor="rgba(0,0,0,0)",
                    layer="above",
                )

            # Stronger emboss (bevel illusion): light on top/left, shadow on bottom/right.
            bevel_inset = inset_step * 0.45
            bx0 = x0 + bevel_inset
            bx1 = x1 - bevel_inset
            by0 = y0 + bevel_inset
            by1 = y1 - bevel_inset

            bevel_w = max(3.0, float(cfg.core_block_border_width) * 1.9)

            # top (light)
            fig.add_shape(
                type="line",
                x0=bx0, y0=by1,
                x1=bx1, y1=by1,
                line=dict(color="rgba(255,255,255,1.0)", width=bevel_w),
                layer="above",
            )
            # left (light)
            fig.add_shape(
                type="line",
                x0=bx0, y0=by0,
                x1=bx0, y1=by1,
                line=dict(color="rgba(255,255,255,1.0)", width=bevel_w),
                layer="above",
            )
            # bottom (shadow)
            fig.add_shape(
                type="line",
                x0=bx0, y0=by0,
                x1=bx1, y1=by0,
                line=dict(color="rgba(0,0,0,0.75)", width=bevel_w),
                layer="above",
            )
            # right (shadow)
            fig.add_shape(
                type="line",
                x0=bx1, y0=by0,
                x1=bx1, y1=by1,
                line=dict(color="rgba(0,0,0,0.75)", width=bevel_w),
                layer="above",
            )

        if isinstance(active_id, str) and active_id.startswith("BLOCK::"):
            # Locate the CORE block trace bbox and apply embossed+glow overlays.
            for idx2, tr2 in enumerate(fig.data):
                if idx2 >= len(curve_to_id):
                    break

                oid2 = curve_to_id[idx2]
                if oid2.endswith("::TEXT"):
                    continue
                if oid2 != active_id:
                    continue
                if is_block_hitbox_trace(tr2):
                    continue

                fc2 = getattr(tr2, "fillcolor", None)
                if (fc2 is None) or (fc2 == "rgba(0,0,0,0)"):
                    continue

                try:
                    xs2 = list(tr2.x) if tr2.x is not None else []
                    ys2 = list(tr2.y) if tr2.y is not None else []
                except Exception:
                    continue

                if not xs2 or not ys2:
                    continue

                _add_core_rect_fx(min(xs2), max(xs2), min(ys2), max(ys2))
                break

        elif isinstance(active_id, str) and active_id.startswith("BAND::"):
            band_shape_indices = meta.get("band_shape_indices", {}) or {}
            si = band_shape_indices.get(active_id, None)
            if (si is not None) and hasattr(fig.layout, "shapes") and fig.layout.shapes:
                try:
                    si_int = int(si)
                except Exception:
                    si_int = None

                if (si_int is not None) and (0 <= si_int < len(fig.layout.shapes)):
                    shp = fig.layout.shapes[si_int]
                    p = str(getattr(shp, "path", "") or "")
                    band_color = str(getattr(shp, "fillcolor", "") or "rgba(150,150,200,1)")
                    if p:
                        # Parse polygon vertices from the SVG path string.
                        # _poly_to_path produces: M x0,y0 L x1,y1 ... L x_close,y_close Z
                        # Existence mode: 5 coord-pairs (4 vertices + close)
                        # Strength mode:  45 coord-pairs (44 vertices + close)
                        # Polygon winding: p0=bot-left, p1=bot-right, ..right-side..,
                        #   p[half]=top-right, p[half+1]=top-left, ..left-side.., p[-1]=bot-left(close)
                        # Left border:  p[half+1] → ... → p[-1]  (top-left down to bot-left)
                        # Right border: p[1] → ... → p[half]     (bot-right up to top-right)
                        import re as _re
                        try:
                            nums = _re.findall(
                                r'[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?', p)
                            n_pts = len(nums) // 2
                            pts = [(float(nums[i * 2]), float(nums[i * 2 + 1]))
                                   for i in range(n_pts)]
                            # pts[-1] is the explicit close point (== pts[0])
                            real = pts[:-1]        # strip close → N real vertices
                            N = len(real)
                            half = N // 2          # existence=2, strength=22

                            def _pts_to_path(ps):
                                if not ps:
                                    return ""
                                parts = [f"M {ps[0][0]},{ps[0][1]}"]
                                for x, y in ps[1:]:
                                    parts.append(f"L {x},{y}")
                                return " ".join(parts)

                            # Right edge: real[1..half]
                            right_path = _pts_to_path(real[1:half + 1])
                            # Left edge: real[half+1..N-1] then close to real[0]
                            left_path  = _pts_to_path(real[half + 1:] + [real[0]])

                            band_layer = getattr(shp, "layer", "below")
                            line_kw = dict(color="rgba(0,0,0,0.55)", width=3)
                            for edge_path in (left_path, right_path):
                                if edge_path:
                                    fig.add_shape(
                                        type="path", path=edge_path,
                                        line=line_kw,
                                        fillcolor="rgba(0,0,0,0)",
                                        layer=band_layer,
                                    )
                        except Exception:
                            pass
    except Exception:
        pass

    return fig.to_dict()


def apply_exo_visibility(fig_json: dict, meta: dict, hidden_slices: List[str]) -> Tuple[dict, Set[str]]:
    """
    Returns: (modified_fig_dict, set_of_hidden_band_ids)
    """
    curve_to_id = meta.get("curve_to_id", [])
    block_to_bands = meta.get("block_to_bands", {})
    band_shape_indices = meta.get("band_shape_indices", {}) or {}
    block_text_anno_indices = meta.get("block_text_anno_indices", {}) or {}
    slices = meta.get("slices", [])

    hidden = set(hidden_slices) & set(slices)
    hide_ids: Set[str] = set()
    hidden_bands: Set[str] = set()
    for sid in hidden:
        exo_block = f"BLOCK::{sid}::{EXO_NAME}"
        hide_ids.add(exo_block)
        for band in block_to_bands.get(exo_block, []):
            hide_ids.add(band)
            hidden_bands.add(band)

    for idx, tr in enumerate(fig_json.get("data", [])):
        if idx >= len(curve_to_id):
            continue

        oid = curve_to_id[idx]

        if oid == "BG::CLICK":
            tr["visible"] = True
            continue

        if oid.startswith("BTNEXO::"):
            sid = oid.split("::", 1)[1]
            tr["visible"] = (sid in hidden)
            continue

        base_oid = oid.replace("::TEXT", "") if oid.endswith("::TEXT") else oid
        if base_oid in hide_ids:
            tr["visible"] = False
        else:
            tr["visible"] = True

    # Also hide/show band SHAPES (since bands are drawn as shapes, not just traces)
    layout = fig_json.get("layout", {})
    shapes = layout.get("shapes", [])
    if band_shape_indices and shapes:
        for band_oid, si in band_shape_indices.items():
            try:
                si_int = int(si)
            except Exception:
                continue
            if si_int < 0 or si_int >= len(shapes):
                continue

            # Hide this band shape if it's connected to a hidden Exo
            if str(band_oid) in hide_ids:
                shapes[si_int]["visible"] = False
            else:
                shapes[si_int]["visible"] = True

    # Also hide/show block text ANNOTATIONS (e.g., "Exo" text when Exo block is hidden)
    annotations = layout.get("annotations", [])
    if block_text_anno_indices and annotations:
        for block_oid, aidx in block_text_anno_indices.items():
            try:
                aidx_int = int(aidx)
            except Exception:
                continue
            if aidx_int < 0 or aidx_int >= len(annotations):
                continue

            # Hide this block text annotation if its block is hidden
            if str(block_oid) in hide_ids:
                annotations[aidx_int]["visible"] = False
            else:
                annotations[aidx_int]["visible"] = True

    return fig_json, hidden_bands


def apply_band_visibility(fig_json: dict, meta: dict, band_state: dict, exo_hidden_bands: Set[str] = None) -> dict:
    """
    Hide/show bands by TYPE (primary label), affecting both:
      - trace-based hover catchers (fig.data)
      - shape-based band polygons (fig.layout.shapes)

    band_state format:
      {"hidden_types": ["Inflow", "Outflow", ...]}
      Special: "__ALL__" means hide all types.
    """
    curve_to_id = meta.get("curve_to_id", [])
    band_shape_indices = meta.get("band_shape_indices", {}) or {}
    band_primary_by_oid = meta.get("band_primary_by_oid", {}) or {}
    if exo_hidden_bands is None:
        exo_hidden_bands = set()

    hidden_types = set()
    if isinstance(band_state, dict):
        hidden_types = set([str(x) for x in (band_state.get("hidden_types", []) or [])])

    hide_all = ("__ALL__" in hidden_types)

    def _should_hide_band(band_oid: str) -> bool:
        if hide_all:
            return True
        primary = str(band_primary_by_oid.get(str(band_oid), "Unknown"))
        return (primary in hidden_types)

    # 1) Toggle band traces (includes hover catchers)
    for idx, tr in enumerate(fig_json.get("data", [])):
        if idx >= len(curve_to_id):
            continue

        oid = curve_to_id[idx]
        base_oid = oid.replace("::TEXT", "") if oid.endswith("::TEXT") else oid

        if base_oid.startswith("BAND::"):
            if base_oid in exo_hidden_bands:
                tr["visible"] = False
            else:
                tr["visible"] = (not _should_hide_band(base_oid))

    # 2) Toggle band shapes (polygons)
    def _as_index_list(v: Any) -> List[Any]:
        if v is None:
            return []
        if isinstance(v, (list, tuple, set)):
            return list(v)
        return [v]

    layout = fig_json.get("layout", {})
    shapes = layout.get("shapes", [])
    if shapes:
        for boid, idx_list in band_shape_indices.items():
            for si in _as_index_list(idx_list):
                try:
                    si_int = int(si)
                except Exception:
                    continue
                if 0 <= si_int < len(shapes):
                    if str(boid) in exo_hidden_bands:
                        shapes[si_int]["visible"] = False
                    else:
                        shapes[si_int]["visible"] = (not _should_hide_band(str(boid)))

    return fig_json


# --- Module-level text helpers (used by both build_figure and apply_enclosure_label_visibility) ---
def _is_cjk_mod(ch):
    cp = ord(ch)
    return (0x4E00 <= cp <= 0x9FFF or 0x3400 <= cp <= 0x4DBF or
            0x2E80 <= cp <= 0x2EFF or 0x3000 <= cp <= 0x303F or 0xFF00 <= cp <= 0xFFEF)


def _visual_len_mod(s):
    return sum(1.7 if _is_cjk_mod(ch) else 1.0 for ch in s)


# Module-level Arial em-width table (mirrors _ARIAL_EM in build_figure).
# Used by _wrap_text_mod's pixel-accurate mode for enclosure / aggregation
# name wrapping. See the in-build_figure copy for the metric source and
# rationale (Adobe Arial AFM).
_ARIAL_EM_MOD = {
    ' ': 0.278, '!': 0.278, '"': 0.355, '#': 0.556, '$': 0.556, '%': 0.889,
    '&': 0.667, "'": 0.191, '(': 0.333, ')': 0.333, '*': 0.389, '+': 0.584,
    ',': 0.278, '-': 0.333, '.': 0.278, '/': 0.278,
    '0': 0.556, '1': 0.556, '2': 0.556, '3': 0.556, '4': 0.556,
    '5': 0.556, '6': 0.556, '7': 0.556, '8': 0.556, '9': 0.556,
    ':': 0.278, ';': 0.278, '<': 0.584, '=': 0.584, '>': 0.584, '?': 0.556,
    '@': 1.015,
    'A': 0.667, 'B': 0.667, 'C': 0.722, 'D': 0.722, 'E': 0.667, 'F': 0.611,
    'G': 0.778, 'H': 0.722, 'I': 0.278, 'J': 0.500, 'K': 0.667, 'L': 0.556,
    'M': 0.833, 'N': 0.722, 'O': 0.778, 'P': 0.667, 'Q': 0.778, 'R': 0.722,
    'S': 0.667, 'T': 0.611, 'U': 0.722, 'V': 0.667, 'W': 0.944, 'X': 0.667,
    'Y': 0.667, 'Z': 0.611,
    '[': 0.278, '\\': 0.278, ']': 0.278, '^': 0.469, '_': 0.556, '`': 0.333,
    'a': 0.556, 'b': 0.556, 'c': 0.500, 'd': 0.556, 'e': 0.556, 'f': 0.278,
    'g': 0.556, 'h': 0.556, 'i': 0.222, 'j': 0.222, 'k': 0.500, 'l': 0.222,
    'm': 0.833, 'n': 0.556, 'o': 0.556, 'p': 0.556, 'q': 0.556, 'r': 0.333,
    's': 0.500, 't': 0.278, 'u': 0.556, 'v': 0.500, 'w': 0.722, 'x': 0.500,
    'y': 0.500, 'z': 0.500,
    '{': 0.334, '|': 0.260, '}': 0.334, '~': 0.584,
}
_ARIAL_EM_MOD_DEFAULT = 0.556
_CJK_EM_MOD = 1.0


def _char_em_mod(ch):
    if _is_cjk_mod(ch):
        return _CJK_EM_MOD
    return _ARIAL_EM_MOD.get(ch, _ARIAL_EM_MOD_DEFAULT)


def _text_px_width_mod(text, font_size):
    return sum(_char_em_mod(c) for c in text) * font_size


def _wrap_text_mod(text, max_chars, max_lines_limit, *, font_size=None):
    """Wrap text into lines.

    Two measurement modes (mirrors _wrap_text in build_figure):
      • PIXEL MODE  (font_size given AND > 0):
        `max_chars` is treated as a PIXEL budget (max line width).
        Each character's real Arial em-width is accumulated as we walk
        the string; the moment cumulative width exceeds the budget we
        break. Wide letters (M, W) push past the wall sooner than narrow
        letters (i, l, j) — guaranteeing zero overflow regardless of the
        text's character composition. This is the "wall-bump → break"
        behaviour applied to enclosure / aggregation name labels.
      • CHARACTER MODE  (font_size None — legacy):
        `max_chars` is a character count; CJK = 1.7 each, others = 1.0.

    Hyphenation: when no natural separator is near the break point AND
    we're splitting an English word mid-character, a hyphen is appended.
    """
    # Mode-specific measurement helpers.
    if font_size is not None and font_size > 0:
        def _ch_w(ch):
            return _char_em_mod(ch) * font_size
        def _measure(s):
            return _text_px_width_mod(s, font_size)
    else:
        def _ch_w(ch):
            return 1.7 if _is_cjk_mod(ch) else 1.0
        def _measure(s):
            return _visual_len_mod(s)
    budget = max_chars

    if _measure(text) <= budget:
        return [text]

    def _is_ascii_letter(ch):
        return ch.isalpha() and not _is_cjk_mod(ch)

    lines = []
    remaining = text
    last_line_text_start = 0
    while remaining and len(lines) < max_lines_limit:
        last_line_text_start = len(text) - len(remaining)
        if _measure(remaining) <= budget:
            lines.append(remaining)
            remaining = ""
            break
        break_point = 0
        vlen = 0.0
        for ci, ch in enumerate(remaining):
            vlen += _ch_w(ch)
            if vlen > budget:
                break
            break_point = ci + 1
        if break_point == 0:
            break_point = 1
        best_sep = -1
        min_pos = max(1, break_point // 3)
        for si in range(break_point - 1, min_pos - 1, -1):
            ch = remaining[si]
            if ch in (' ', ',', '_', '-', '(', ')', '/', '、', '，'):
                best_sep = si + 1
                break
            if si > 0 and (_is_cjk_mod(remaining[si]) != _is_cjk_mod(remaining[si - 1])):
                best_sep = si
                break
        if best_sep > 0:
            break_point = best_sep
            line = remaining[:break_point].rstrip()
        else:
            # Mid-word break — try to hyphenate if both neighbours are ASCII letters.
            min_budget_for_hyphen = 3 if font_size is None else 3 * _ARIAL_EM_MOD_DEFAULT * font_size
            can_hyphenate = (
                budget >= min_budget_for_hyphen
                and break_point >= 2
                and break_point < len(remaining)
                and _is_ascii_letter(remaining[break_point - 1])
                and _is_ascii_letter(remaining[break_point])
            )
            if can_hyphenate:
                hyph_break = break_point - 1
                prefix = remaining[:hyph_break].rstrip()
                if prefix:
                    line = prefix + "-"
                    break_point = hyph_break
                else:
                    line = remaining[:break_point].rstrip()
            else:
                line = remaining[:break_point].rstrip()
        # Hard-cap fallback (should rarely trigger thanks to hyphenation above).
        if _measure(line) > budget:
            while line and _measure(line + "…") > budget:
                line = line[:-1]
            line = (line + "…") if line else "…"
        lines.append(line)
        remaining = remaining[break_point:].lstrip()
    if remaining and len(lines) == max_lines_limit:
        last_line = lines[-1]
        if _measure(last_line) + _measure(remaining) > budget:
            # Truncate the last line with an ellipsis. (See full
            # explanatory comment in build_figure._wrap_text for the
            # "-" vs "…" semantic distinction.)
            if last_line.endswith("-"):
                last_line = last_line[:-1]
            while last_line and _measure(last_line + "…") > budget:
                last_line = last_line[:-1]
            lines[-1] = (last_line + "…") if last_line else "…"
    return lines if lines else [text]


def apply_enclosure_label_visibility(fig_json: dict, meta: dict, level_state: dict, cfg: StyleConfig,
                                     style_state: dict = None, graph_size: dict = None) -> dict:
    """
    当group名字显示时：
    - 将所有元素（blocks, bands, enclosures）向右移动，为group名字腾出空间
    - 这样可以保持所有组件之间的相对距离不变
    - 应用label样式设置（字体、字号、颜色）
    """
    style_state = style_state or {}
    _gs_w = (graph_size or {}).get("w") if graph_size else None
    _gs_h = (graph_size or {}).get("h") if graph_size else None
    _CANVAS_W = float(_gs_w) if (_gs_w and _gs_w > 100) else 600.0
    _CANVAS_H = float(_gs_h) if (_gs_h and _gs_h > 100) else 700.0
    show_map = {}
    if isinstance(level_state, dict):
        show_map = dict(level_state.get("show", {}) or {})

    level_keys = list(meta.get("enclosure_label_level_keys", []))
    mask_indices = list(meta.get("enclosure_label_mask_shape_indices", []))
    border_indices = list(meta.get("enclosure_label_border_shape_indices", []))
    anno_indices = list(meta.get("enclosure_label_anno_indices", []))
    level_depths = list(meta.get("enclosure_label_depths", []))
    slices_list = meta.get("slices", [])
    curve_to_id = meta.get("curve_to_id", [])
    band_shape_indices = meta.get("band_shape_indices", {}) or {}

    fig = go.Figure(fig_json)

    if (not hasattr(fig.layout, "shapes")) or (not fig.layout.shapes):
        return fig.to_dict()

    n = min(len(level_keys), len(mask_indices), len(border_indices), len(anno_indices), len(level_depths))
    if n <= 0:
        return fig.to_dict()

    # === Helper functions ===
    def _safe_float(v: Any, default: float = 0.0) -> float:
        try:
            if v is None or v == "" or v == "None":
                return float(default)
            return float(v)
        except Exception:
            return float(default)

    def _bb_of_shape_idx(si: int) -> Optional[Tuple[float, float, float, float]]:
        if si < 0 or si >= len(fig.layout.shapes):
            return None
        shp = fig.layout.shapes[si]
        rx0 = getattr(shp, "x0", None)
        rx1 = getattr(shp, "x1", None)
        ry0 = getattr(shp, "y0", None)
        ry1 = getattr(shp, "y1", None)
        if (rx0 is None) or (rx1 is None) or (ry0 is None) or (ry1 is None):
            return None
        x0, x1 = _safe_float(rx0, 0.0), _safe_float(rx1, 0.0)
        y0, y1 = _safe_float(ry0, 0.0), _safe_float(ry1, 0.0)
        if x1 < x0:
            x0, x1 = x1, x0
        if y1 < y0:
            y0, y1 = y1, y0
        return x0, x1, y0, y1

    def _contains(outer_bb, inner_bb, eps=1e-6) -> bool:
        ox0, ox1, oy0, oy1 = outer_bb
        ix0, ix1, iy0, iy1 = inner_bb
        return (ox0 - eps <= ix0) and (ix1 <= ox1 + eps) and (oy0 - eps <= iy0) and (iy1 <= oy1 + eps)

    def _get_slice_idx_by_y(y_val: float) -> int:
        layer_step = float(meta.get("y_step", cfg.layer_gap + cfg.block_height))
        for sl_idx in range(len(slices_list)):
            y_center = -sl_idx * layer_step
            if abs(y_val - y_center) < layer_step / 2.0 + 1.0:
                return sl_idx
        return -1

    def _get_enclosure_radii_scaled(level_key: str) -> List[float]:
        """
        Return radii in *data units* for the given enclosure level key ("group", "supergroup1", "supergroup2", ...).
        The UI stores radii as 0..50-like values; we convert them using the same scale factor as shape construction.
        """
        try:
            block_w = float(cfg.block_width)
            block_h = float(cfg.block_height)
        except Exception:
            block_w = float(getattr(cfg, 'block_width', 1.0) or 1.0)
            block_h = float(getattr(cfg, 'block_height', 1.0) or 1.0)
        scale_factor = min(block_w, block_h) / 50.0 if min(block_w, block_h) > 0 else 0.0

        radii_raw = [0, 0, 0, 0]
        if str(level_key) == 'group':
            base = (style_state or {}).get('group', {}) if isinstance(style_state, dict) else {}
            rr = base.get('radii', None)
            if rr is None:
                rr = getattr(cfg, 'group_style', {}).get('radii', [0, 0, 0, 0])
            if isinstance(rr, (list, tuple)):
                radii_raw = list(rr)
        else:
            # Supergroup styles are stored under style_state['supergroups'][<level_str>]
            k_str = str(level_key)
            sg_level_str = k_str.replace('supergroup', '') if k_str.startswith('supergroup') else '1'
            supergroups_style = (style_state or {}).get('supergroups', {}) if isinstance(style_state, dict) else {}
            base = supergroups_style.get(str(sg_level_str), {}) if isinstance(supergroups_style, dict) else {}
            rr = base.get('radii', None)
            if rr is None:
                rr = getattr(cfg, 'supergroup_styles', {}).get(str(sg_level_str), {}).get('radii', [0, 0, 0, 0])
            if isinstance(rr, (list, tuple)):
                radii_raw = list(rr)

        while len(radii_raw) < 4:
            radii_raw.append(0)
        radii_raw = radii_raw[:4]
        return [(r if r is not None else 0) * scale_factor for r in radii_raw]

    def _sync_enclosure_border_path(bidx: int, level_key: str, bb: Tuple[float, float, float, float]) -> None:
        """
        If the enclosure border is a rounded-corner path shape, its path string must be rebuilt
        whenever x0/x1/y0/y1 change (e.g., global_shift, supergroup Y-expansion).
        Otherwise, the path can visually drift relative to its bbox.
        """
        try:
            if bidx < 0 or bidx >= len(fig.layout.shapes):
                return
            shp = fig.layout.shapes[bidx]
            if str(getattr(shp, 'type', '')) != 'path':
                return
            radii_scaled = _get_enclosure_radii_scaled(level_key)
            x0, x1, y0, y1 = float(bb[0]), float(bb[1]), float(bb[2]), float(bb[3])
            shp.path = rounded_rect_svg_path(x0, y0, x1, y1, radii_scaled)
        except Exception:
            pass

    # === Build items ===
    items: List[Dict[str, Any]] = []
    for i in range(n):
        bidx, midx, aidx = int(border_indices[i]), int(mask_indices[i]), int(anno_indices[i])
        if bidx < 0 or bidx >= len(fig.layout.shapes):
            continue
        if midx < 0 or midx >= len(fig.layout.shapes):
            continue
        if not hasattr(fig.layout, "annotations") or fig.layout.annotations is None:
            continue
        if aidx < 0 or aidx >= len(fig.layout.annotations):
            continue

        k, d = str(level_keys[i]), int(level_depths[i])
        bb = _bb_of_shape_idx(bidx)
        if bb is None:
            continue
        y_mid = (bb[2] + bb[3]) / 2.0
        sl_idx = _get_slice_idx_by_y(y_mid)

        txt = str(fig.layout.annotations[aidx].text or "") if aidx < len(fig.layout.annotations) else ""

        items.append({
            "i": i, "k": k, "d": d,
            "bidx": bidx, "midx": midx, "aidx": aidx,
            "orig_bb": bb, "slice_idx": sl_idx, "text": txt,
        })

    if not items:
        return fig.to_dict()

    # === Build parent-child relationships ===
    for inner in items:
        inner_bb, inner_d = inner["orig_bb"], inner["d"]
        parent = None
        parent_d = 999999
        for outer in items:
            if outer["slice_idx"] != inner["slice_idx"]:
                continue
            if outer["d"] <= inner_d:
                continue
            if _contains(outer["orig_bb"], inner_bb):
                if outer["d"] < parent_d:
                    parent = outer
                    parent_d = outer["d"]
        inner["parent"] = parent

    # Build children lists
    for it in items:
        it["children"] = []
    for it in items:
        p = it.get("parent")
        if p is not None:
            p["children"].append(it)

    # === Check if group labels are shown ===
    show_group = bool(show_map.get("group", False))

    # === Calculate global shift amount ===
    # When group labels are shown, we need to shift everything to the right
    # to make room for the labels while maintaining relative distances
    global_shift = 0.0

    group_style = (style_state or {}).get("group", {}) if isinstance(style_state, dict) else {}
    label_offset_x = float(group_style.get("label_offset_x") if group_style.get(
        "label_offset_x") is not None else cfg.enclosure_group_label_offset_x)
    label_size = float(
        group_style.get("label_size") if group_style.get("label_size") is not None else cfg.enclosure_label_text_size)
    base_label_size = float(cfg.enclosure_label_text_size) if float(cfg.enclosure_label_text_size) else 1.0
    size_scale = label_size / base_label_size

    if show_group:
        # Compute label footprint in x-direction, accounting for rotation.
        # Group labels are typically rotated 90° → x-footprint is text HEIGHT, not width.
        _layout = fig_json.get("layout", {}) if isinstance(fig_json, dict) else {}
        _xr = (_layout.get("xaxis") or {}).get("range", [-10, 10])
        _yr = (_layout.get("yaxis") or {}).get("range", [-10, 10])
        try:
            _x_span = abs(float(_xr[1]) - float(_xr[0]))
        except Exception:
            _x_span = 20.0
        try:
            _y_span = abs(float(_yr[1]) - float(_yr[0]))
        except Exception:
            _y_span = 20.0
        _ppu_x = float(_CANVAS_W) * 0.88 / max(_x_span, 0.01)   # full effective px/unit (matches build_figure now that _fit_ppu_* cap is removed)
        _ppu_y = float(_CANVAS_H) * 0.88 / max(_y_span, 0.01)
        # Group labels are rotated 90°: characters run along y (enclosure height),
        # lines stack along x. Use ppu_y for char dimension, ppu_x for x-extent.

        import math as _math_shift
        label_rotation = int(group_style.get("label_rotation", 90))
        _label_line_spacing_ext = float(group_style.get("label_line_spacing", 0))
        _label_pad_b_ext = float(group_style.get("label_pad_b", 0.0))
        _label_pad_t_ext = float(group_style.get("label_pad_t", 0.0))
        angle_rad = _math_shift.radians(label_rotation)

        for it in items:
            if it["d"] != 0:
                continue
            txt = it["text"]
            # Wrap text based on parent enclosure height (not block_height)
            # Use the GROUP's own y-extent for char budget.
            # Using parent_it height was the bug: parent spans many groups → huge avail_px → no truncation.
            encl_h = abs(it["orig_bb"][3] - it["orig_bb"][2])
            encl_h = max(float(cfg.block_height), encl_h)   # never less than one block
            encl_h = max(0.01, encl_h - _label_pad_b_ext - _label_pad_t_ext)
            avail_px_gs = encl_h * _ppu_y   # chars run along y (enclosure height)
            _effective_line_h_gs = max(label_size * 0.3, label_size * 1.2 + float(_label_line_spacing_ext))
            # Lines cap (vertical safety). Width is now decided by pixel
            # measurement inside _wrap_text_mod when font_size is given.
            _max_lines_gs = max(1, int(avail_px_gs / max(_effective_line_h_gs, 1)))
            _max_lines_gs = min(_max_lines_gs, 2)  # Aggregation name: hard limit 2 lines
            wrapped_gs = _wrap_text_mod(
                txt,
                avail_px_gs,                # pixel budget for line width
                _max_lines_gs,
                font_size=label_size,       # enables pixel-accurate measurement
            )
            n_lines_gs = max(1, len(wrapped_gs))
            # Real measured width of the widest line (replaces the old
            # `max_line_chars * label_size * 0.63` estimate). Using the
            # actual rendered width here lines up the bbox-extent
            # calculation below with what _wrap_text_mod just produced.
            text_w_px = max((_text_px_width_mod(line, label_size) for line in wrapped_gs), default=0.0)
            text_h_px = n_lines_gs * _effective_line_h_gs  # total height with line spacing
            # X-projection of rotated text bbox
            x_extent_px = abs(text_w_px * _math_shift.cos(angle_rad)) + abs(text_h_px * _math_shift.sin(angle_rad))
            # Convert to data units using ppu_x (x-direction)
            x_extent_data = x_extent_px / max(_ppu_x, 1.0)
            # Per-group label extension = text width + distance
            label_extension = x_extent_data + float(label_offset_x)
            it["label_extension"] = label_extension
            it["x_extent_data"] = x_extent_data
            if label_extension > global_shift:
                global_shift = label_extension

    # === Build per-group shift map ===
    # Each group gets a CUMULATIVE shift = sum of label_extensions for groups 0..k.
    # This way each name only takes the space it needs.
    # Layout: [encl_pad] [name_0][dist][group_0] [outer_gap] [name_1][dist][group_1] ...

    # Build per-slice group x-boundary info
    group_shift_map: Dict[int, Dict[str, Any]] = {}  # slice_idx -> {boundaries, shifts, max_shift}
    if show_group and global_shift > 0:
        from collections import defaultdict
        groups_by_slice: Dict[int, list] = defaultdict(list)
        for it in items:
            if it["d"] == 0:
                groups_by_slice[it["slice_idx"]].append(it)

        for sl_idx, group_items in groups_by_slice.items():
            # Sort groups by x0 (left to right)
            group_items.sort(key=lambda g: g["orig_bb"][0])
            n_groups = len(group_items)

            # Build boundary intervals with per-group cumulative shifts
            boundaries = []  # list of (x0, x1, shift)
            cumulative_shift = 0.0
            for gk, git in enumerate(group_items):
                gx0, gx1 = git["orig_bb"][0], git["orig_bb"][1]
                cumulative_shift += git.get("label_extension", global_shift)
                shift = cumulative_shift
                boundaries.append((gx0, gx1, shift))
                git["group_shift"] = shift

            group_shift_map[sl_idx] = {
                "boundaries": boundaries,
                "n_groups": n_groups,
                "first_shift": boundaries[0][2] if boundaries else global_shift,
                "last_shift": boundaries[-1][2] if boundaries else global_shift,
            }

    def _get_shift_for_x(x_val: float, y_val: float) -> float:
        """Get the appropriate shift for a given x,y coordinate."""
        if not show_group or global_shift <= 0:
            return 0.0
        # Determine which slice this y belongs to
        layer_step = float(meta.get("y_step", cfg.layer_gap + cfg.block_height))
        best_sl_idx = -1
        best_dist = float("inf")
        for sl_idx in group_shift_map:
            y_center = -sl_idx * layer_step
            dist = abs(y_val - y_center)
            if dist < best_dist:
                best_dist = dist
                best_sl_idx = sl_idx
        if best_sl_idx < 0:
            return global_shift  # fallback: uniform shift

        info = group_shift_map[best_sl_idx]
        boundaries = info["boundaries"]
        if not boundaries:
            return global_shift

        # Check which group this x falls in (or between)
        for gx0, gx1, shift in boundaries:
            if x_val >= gx0 - 0.01 and x_val <= gx1 + 0.01:
                return shift

        # If to the left of all groups (Exo), don't shift (name goes to right of Exo)
        if x_val < boundaries[0][0]:
            return 0.0
        # If to the right of all groups, shift to match enclosure right edge
        if x_val > boundaries[-1][1]:
            return info["last_shift"]

        # Between groups: use the shift of the group to the right
        for gx0, gx1, shift in boundaries:
            if x_val < gx0:
                return shift
        return info["last_shift"]

    # === Apply per-group shift to ALL traces ===
    if show_group and global_shift > 0:
        for idx, tr in enumerate(fig.data):
            try:
                if tr.x is not None and tr.y is not None and len(tr.x) > 0:
                    # Get representative y for this trace
                    y_vals = [y for y in tr.y if y is not None]
                    y_rep = y_vals[0] if y_vals else 0.0
                    # For traces that span multiple groups (bands), shift each point individually
                    new_x = []
                    for xi, x_val in enumerate(tr.x):
                        y_val = float(tr.y[xi]) if xi < len(tr.y) and tr.y[xi] is not None else y_rep
                        new_x.append(x_val + _get_shift_for_x(x_val, y_val))
                    tr.x = new_x
            except Exception:
                pass
    elif global_shift > 0:
        pass  # no shift when group names hidden

    # === Apply per-group shift to ALL annotations ===
    enclosure_anno_indices_set = set(anno_indices)
    if show_group and global_shift > 0 and hasattr(fig.layout, "annotations") and fig.layout.annotations:
        for aidx, ann in enumerate(fig.layout.annotations):
            try:
                current_x = float(getattr(ann, "x", 0))
                current_y = float(getattr(ann, "y", 0))
                ann.x = current_x + _get_shift_for_x(current_x, current_y)
            except Exception:
                pass

    # === Apply per-group shift to ALL band shapes ===
    # Each band connects a source block to a dest block. To shift correctly:
    # 1) Parse band OID → src_slice, src_cat, dst_slice, dst_cat
    # 2) Look up each block's center-x from meta["block_cx"]
    # 3) Use _get_shift_for_x(block_center_x, layer_y_center) to get each end's group shift
    # 4) Linearly interpolate shift by y-fraction across the band's height
    if show_group and global_shift > 0 and band_shape_indices:
        _band_layer_step = float(meta.get("y_step", cfg.layer_gap + cfg.block_height))
        _band_slice_ids = meta.get("slices", [])
        _band_sid_to_idx = {str(sid): idx for idx, sid in enumerate(_band_slice_ids)}
        _band_block_cx = meta.get("block_cx", {}) or {}

        for band_oid, si in band_shape_indices.items():
            try:
                si_int = int(si)
                if si_int < 0 or si_int >= len(fig.layout.shapes):
                    continue
                shp = fig.layout.shapes[si_int]
                path_str = getattr(shp, "path", "")
                if not path_str:
                    continue

                # Parse band OID: "BAND::src_slice::src_cat::dst_slice::dst_cat"
                parts = str(band_oid).split("::")
                if len(parts) < 5:
                    continue
                src_slice, src_cat = parts[1], parts[2]
                dst_slice, dst_cat = parts[3], parts[4]

                src_sl_idx = _band_sid_to_idx.get(src_slice, -1)
                dst_sl_idx = _band_sid_to_idx.get(dst_slice, -1)

                # Look up block center-x from meta
                src_cx = _band_block_cx.get(f"{src_slice}::{src_cat}")
                dst_cx = _band_block_cx.get(f"{dst_slice}::{dst_cat}")

                # Compute y-centers for each layer
                src_y_center = -src_sl_idx * _band_layer_step if src_sl_idx >= 0 else 0.0
                dst_y_center = -dst_sl_idx * _band_layer_step if dst_sl_idx >= 0 else 0.0

                # Get group shift for each end using block center-x and layer y-center
                if src_cx is not None and src_sl_idx >= 0:
                    src_shift = _get_shift_for_x(float(src_cx), src_y_center)
                else:
                    src_shift = global_shift  # fallback
                if dst_cx is not None and dst_sl_idx >= 0:
                    dst_shift = _get_shift_for_x(float(dst_cx), dst_y_center)
                else:
                    dst_shift = global_shift  # fallback

                # If both ends get the same shift, apply uniformly (most common case)
                if abs(src_shift - dst_shift) < 1e-9:
                    _s = src_shift

                    def _shift_uniform(match, _s=_s):
                        return f"{match.group(1)} {float(match.group(2)) + _s},{match.group(3)}"

                    shp.path = re.sub(r'([ML])\s*([-\d.]+),([-\d.]+)', _shift_uniform, path_str)
                else:
                    # Cross-group band: interpolate by y-fraction
                    # Get y range from path
                    y_vals = [float(m.group(1)) for m in re.finditer(r'[ML]\s*[-\d.]+,([-\d.]+)', path_str)]
                    if not y_vals:
                        continue
                    y_min_b, y_max_b = min(y_vals), max(y_vals)
                    y_span_b = y_max_b - y_min_b

                    # Determine which end is at y_max (top) and which at y_min (bottom)
                    # Source is the layer with higher y (closer to 0)
                    if src_y_center >= dst_y_center:
                        top_shift, bot_shift = src_shift, dst_shift
                    else:
                        top_shift, bot_shift = dst_shift, src_shift

                    def _shift_interp(match, _ts=top_shift, _bs=bot_shift,
                                      _ymin=y_min_b, _yspan=y_span_b):
                        cmd = match.group(1)
                        x = float(match.group(2))
                        y = float(match.group(3))
                        if _yspan > 1e-6:
                            t = (y - _ymin) / _yspan  # 0 at bottom, 1 at top
                            shift = _ts * t + _bs * (1.0 - t)
                        else:
                            shift = _ts
                        return f"{cmd} {x + shift},{y}"

                    shp.path = re.sub(r'([ML])\s*([-\d.]+),([-\d.]+)', _shift_interp, path_str)
            except Exception:
                pass

    # === Apply shift to enclosure shapes and calculate new positions ===
    # Each group shifts right by its cumulative label_extension.
    # The name sits to the LEFT of the group border.
    # Supergroups: left edge shifts to encompass leftmost child name, right follows rightmost child.
    new_x0: Dict[int, float] = {}
    new_x1: Dict[int, float] = {}

    for it in items:
        orig_x0, orig_x1 = it["orig_bb"][0], it["orig_bb"][1]

        if show_group and global_shift > 0:
            if it["d"] == 0:
                # Group (innermost): use its computed group_shift
                gs = it.get("group_shift", global_shift)
                new_x0[it["i"]] = orig_x0 + gs
                new_x1[it["i"]] = orig_x1 + gs
            else:
                # Supergroup: must encompass ALL shifted children + their names.
                sl_idx = it["slice_idx"]
                info = group_shift_map.get(sl_idx)
                if info and info["boundaries"]:
                    # Find the child groups contained within this supergroup's orig bbox
                    min_shift = float("inf")
                    max_shift = float("-inf")
                    min_child_label_ext = float("inf")
                    for child in items:
                        if child["d"] != 0:
                            continue
                        if child["slice_idx"] != sl_idx:
                            continue
                        cx0 = child["orig_bb"][0]
                        cx1 = child["orig_bb"][1]
                        if cx0 >= orig_x0 - 0.01 and cx1 <= orig_x1 + 0.01:
                            cs = child.get("group_shift", global_shift)
                            cl = child.get("label_extension", global_shift)
                            if cs < min_shift:
                                min_shift = cs
                                min_child_label_ext = cl
                            max_shift = max(max_shift, cs)
                    if min_shift == float("inf"):
                        min_shift = max_shift = global_shift
                        min_child_label_ext = global_shift
                    # Left: leftmost child shifted by min_shift, its name extends
                    # min_child_label_ext to the left of the shifted group border.
                    # So the name left edge = child_orig_x0 + min_shift - min_child_label_ext
                    # SG left must shift by (min_shift - min_child_label_ext) to maintain pad_left.
                    # For leftmost SG: min_shift == min_child_label_ext → shift = 0 ✓
                    # For non-leftmost SG: shift = cumulative of all groups before leftmost child ✓
                    new_x0[it["i"]] = orig_x0 + max(0.0, min_shift - min_child_label_ext)
                    # Right: follow the rightmost group's shift
                    new_x1[it["i"]] = orig_x1 + max_shift
                else:
                    # Fallback: left stays, right expands
                    new_x0[it["i"]] = orig_x0
                    new_x1[it["i"]] = orig_x1 + global_shift
        else:
            new_x0[it["i"]] = orig_x0
            new_x1[it["i"]] = orig_x1

    # === Apply new positions to enclosure border shapes ===
    for it in items:
        bidx = it["bidx"]
        ii = it["i"]

        try:
            fig.layout.shapes[bidx].x0 = new_x0[ii]
            fig.layout.shapes[bidx].x1 = new_x1[ii]
        except Exception:
            pass

        it["new_bb"] = (new_x0[ii], new_x1[ii], it["orig_bb"][2], it["orig_bb"][3])
        _sync_enclosure_border_path(bidx, it["k"], it["new_bb"])

    # === Handle Y expansion for supergroup labels (top-center) ===
    ratio = float(cfg.enclosure_label_text_size) / float(cfg.block_text_size) if cfg.block_text_size > 0 else 0.75
    label_h = float(cfg.block_height) * ratio * float(cfg.enclosure_label_height_scale)
    label_h_min = float(cfg.block_height) * float(cfg.enclosure_label_height_min_ratio)
    if label_h < label_h_min:
        label_h = label_h_min
    half_h = 0.5 * label_h
    down = max(float(cfg.enclosure_label_mask_down_y), half_h)
    up = max(float(cfg.enclosure_label_mask_up_y), half_h)
    clearance = float(cfg.enclosure_label_clearance_y)

    y1_new: Dict[int, float] = {it["i"]: it["new_bb"][3] for it in items}

    items_sorted = sorted(items, key=lambda t: t["d"])
    for inner in items_sorted:
        ik = inner["k"]
        if not bool(show_map.get(ik, False)):
            continue
        if ik == "group":
            continue

        ii = inner["i"]

        sg_level_str = ik.replace("supergroup", "") if ik.startswith("supergroup") else "1"
        supergroups_style = (style_state or {}).get("supergroups", {}) if isinstance(style_state, dict) else {}
        sg_style = supergroups_style.get(sg_level_str, {}) if isinstance(supergroups_style, dict) else {}
        default_sg_label_distance = -0.05
        sg_label_distance = _safe_float(sg_style.get("label_distance"), default_sg_label_distance)

        # Supergroup labels use yanchor="bottom": y is the label bottom edge.
        # Distance is measured in data units (block_height multiples) from the enclosure top border (y1) to label bottom.
        default_sg_label_size_px = 8 if str(sg_level_str) in ("1", "2") else int(cfg.enclosure_label_text_size)
        try:
            sg_label_size_px = int(sg_style.get("label_size", default_sg_label_size_px))
        except Exception:
            sg_label_size_px = int(default_sg_label_size_px)

        # Use scaleanchor-aware conversion for label height
        _layout_ye = fig_json.get("layout", {}) if isinstance(fig_json, dict) else {}
        _xr_ye = (_layout_ye.get("xaxis") or {}).get("range", [-10, 10])
        _yr_ye = (_layout_ye.get("yaxis") or {}).get("range", [-10, 10])
        _x_span_ye = abs(float(_xr_ye[1]) - float(_xr_ye[0]))
        _y_span_ye = abs(float(_yr_ye[1]) - float(_yr_ye[0]))
        _ppu_ye = min(float(_CANVAS_W) * 0.88 / max(_x_span_ye, 0.01),
                      float(_CANVAS_H) * 0.88 / max(_y_span_ye, 0.01))   # full effective px/unit
        label_h_sg = (sg_label_size_px * 1.2 + 4.0) / max(_ppu_ye, 1.0)

        y_anchor_bottom = y1_new[ii] + sg_label_distance * float(cfg.block_height)
        inner_top = y_anchor_bottom + label_h_sg

        parent = inner.get("parent")
        while parent is not None:
            ok = parent["k"]
            oi = parent["i"]
            if bool(show_map.get(ok, False)):
                parent_depth = 0 if ok == "group" else int(str(ok).replace("supergroup", ""))
                parent_down_factor = 0.10 if int(parent_depth) == 0 else 0.05
                outer_down = float(cfg.enclosure_label_mask_down_y) * float(parent_down_factor)
            else:
                outer_down = 0.0
            # Supergroup2 pad_top should control the distance between:
            #   - parent top border and child (inner) supergroup name top, when the child name is shown;
            # while keeping the previous behaviour as the default initial state (pad_top=0.0).
            default_pad_top_abs = 0.0
            parent_pad_top_abs = default_pad_top_abs

            if isinstance(ok, str) and ok.startswith("supergroup"):
                parent_level_str = ok.replace("supergroup", "")
                supergroups_style_parent = (style_state or {}).get("supergroups", {}) if isinstance(style_state,
                                                                                                    dict) else {}
                parent_style = supergroups_style_parent.get(str(parent_level_str), {}) if isinstance(
                    supergroups_style_parent, dict) else {}
                parent_pad_top_abs = _safe_float(parent_style.get("pad_top"), default_pad_top_abs)

            pad_top_delta = parent_pad_top_abs - default_pad_top_abs
            need_outer_y1 = inner_top + clearance + outer_down + pad_top_delta
            if need_outer_y1 > y1_new[oi]:
                y1_new[oi] = need_outer_y1
            parent = parent.get("parent")

    # Apply y1 updates
    for it in items:
        ii = it["i"]
        bidx = it["bidx"]
        try:
            fig.layout.shapes[bidx].y1 = y1_new[ii]
        except Exception:
            pass
        it["new_bb"] = (it["new_bb"][0], it["new_bb"][1], it["new_bb"][2], y1_new[ii])
        _sync_enclosure_border_path(bidx, it["k"], it["new_bb"])

    # === Set annotation visibility and position ===
    for it in items:
        k = it["k"]
        vis = bool(show_map.get(k, False))
        aidx = it["aidx"]
        midx = it["midx"]
        new_bb_x0, new_bb_x1, new_bb_y0, new_bb_y1 = it["new_bb"]

        if not vis:
            try:
                fig.layout.annotations[aidx].visible = False
            except Exception:
                pass
            try:
                fig.layout.shapes[midx].visible = False
            except Exception:
                pass
            continue

        txt = it["text"]

        if k == "group":
            # LEFT side label - positioned in the space between parent enclosure left and group frame left
            group_x0 = new_bb_x0  # shifted group frame x0
            y_center = (new_bb_y0 + new_bb_y1) / 2.0

            # Get group label style settings
            group_style = (style_state or {}).get("group", {}) if isinstance(style_state, dict) else {}
            label_offset_x = float(group_style.get("label_offset_x") if group_style.get(
                "label_offset_x") is not None else cfg.enclosure_group_label_offset_x)
            label_font = group_style.get("label_font", "Arial")
            label_size = int(group_style.get("label_size", 23))
            label_color = group_style.get("label_color", "#000000")
            label_rotation = int(group_style.get("label_rotation", 90))
            label_line_spacing = float(group_style.get("label_line_spacing", 0))
            label_pad_b = float(group_style.get("label_pad_b", 0.0))
            label_pad_t = float(group_style.get("label_pad_t", 0.0))

            # Position: after rotation, text is CENTERED on x regardless of xanchor.
            # So we need: right_visual_edge = group_x0 - label_offset_x
            # right_visual_edge = x + x_extent/2, therefore x = group_x0 - label_offset_x - x_extent/2
            # For rotation=0, xanchor="right" works correctly, so no adjustment needed.
            import math as _math_pos
            _rot_rad = _math_pos.radians(label_rotation)
            _x_ext = it.get("x_extent_data", 0.0)
            if abs(label_rotation) > 1:
                # Rotated: text centers on x after rotation
                x_label = group_x0 - label_offset_x - _x_ext / 2.0
            else:
                # No rotation: xanchor="right" positions right edge at x
                x_label = group_x0 - label_offset_x

            # Wrap text based on parent enclosure height (not block_height)
            # Apply Name padding (B/T) to shrink available space for text
            try:
                _layout_wrap = fig_json.get("layout", {}) if isinstance(fig_json, dict) else {}
                _xr_wrap = (_layout_wrap.get("xaxis") or {}).get("range", [-10, 10])
                _yr_wrap = (_layout_wrap.get("yaxis") or {}).get("range", [-10, 10])
                _x_span_wrap = abs(float(_xr_wrap[1]) - float(_xr_wrap[0]))
                _y_span_wrap = abs(float(_yr_wrap[1]) - float(_yr_wrap[0]))
                _ppu_x_wrap = float(_CANVAS_W) * 0.88 / max(_x_span_wrap, 0.01)   # full effective px/unit
                _ppu_y_wrap = float(_CANVAS_H) * 0.88 / max(_y_span_wrap, 0.01)

                # Available length = parent enclosure height (label is rotated, runs along y)
                # Use the GROUP's own y-extent (same as pass 1).
                # parent_it height was causing max_chars >> actual block space → overflow.
                encl_h = abs(it["orig_bb"][3] - it["orig_bb"][2])
                encl_h = max(float(cfg.block_height), encl_h)
                # Shrink available height by Name padding B+T
                encl_h = max(0.01, encl_h - label_pad_b - label_pad_t)
                avail_px = encl_h * _ppu_y_wrap   # chars along y-axis
                # Account for line_spacing when computing max lines
                _effective_line_h = max(label_size * 0.3, label_size * 1.2 + float(label_line_spacing))
                # Lines safety cap (vertical packing). Pixel-accurate width
                # is handled by _wrap_text_mod itself when font_size is given,
                # so we no longer need the 0.72 char-width estimate; pass the
                # raw pixel budget as `max_chars` (it's reused as the budget).
                _max_lines_fit = max(1, int(avail_px / max(_effective_line_h, 1)) - 1)
                _max_lines_fit = min(_max_lines_fit, 2)  # Aggregation name: hard limit 2 lines
                wrapped = _wrap_text_mod(
                    txt,
                    avail_px,                  # pixel budget for line width
                    _max_lines_fit,
                    font_size=label_size,      # enables pixel-accurate measurement
                )
                # Build display text with line-spacing-aware separators
                # Plotly annotations don't support CSS line-height; simulate with extra <br> tags
                _extra_brs = ""
                _agg_base_lh = label_size * 1.2
                if label_line_spacing > 0 and _agg_base_lh > 0:
                    _n_extra = max(0, round(float(label_line_spacing) / _agg_base_lh))
                    if _n_extra > 0:
                        _extra_brs = "<br>" * _n_extra
                display_txt = ("<br>" + _extra_brs).join(wrapped)
            except Exception:
                display_txt = txt

            # Compute y_center offset from Name padding (asymmetric B/T shifts the center)
            # Shift is in data units: positive pad_b pushes text up, positive pad_t pushes text down
            _pad_shift_y = (label_pad_b - label_pad_t) / 2.0
            # For rotated labels, the y-axis of the annotation aligns with the label's main axis
            if abs(label_rotation) > 45:
                y_label = y_center + _pad_shift_y
            else:
                y_label = y_center + _pad_shift_y

            try:
                fig.layout.annotations[aidx].visible = True
                fig.layout.annotations[aidx].x = x_label
                fig.layout.annotations[aidx].y = y_label
                fig.layout.annotations[aidx].text = display_txt
                fig.layout.annotations[aidx].xanchor = "center" if abs(label_rotation) > 1 else "right"
                fig.layout.annotations[aidx].yanchor = "middle"
                fig.layout.annotations[aidx].align = "center"
                fig.layout.annotations[aidx].bgcolor = "rgba(0,0,0,0)"
                fig.layout.annotations[aidx].bordercolor = "rgba(0,0,0,0)"
                fig.layout.annotations[aidx].borderpad = 0
                fig.layout.annotations[aidx].textangle = -label_rotation
                fig.layout.annotations[aidx].font = dict(
                    family=label_font,
                    size=label_size,
                    color=label_color
                )
            except Exception:
                pass

            # Hide the white background mask (no longer needed)
            try:
                fig.layout.shapes[midx].visible = False
            except Exception:
                pass


        else:

            # TOP center label (supergroup)

            xmid = (new_bb_x0 + new_bb_x1) / 2.0

            # Get supergroup label style settings
            # k is like "supergroup1", "supergroup2", etc.
            sg_level_str = k.replace("supergroup", "") if k.startswith("supergroup") else "1"
            supergroups_style = style_state.get("supergroups", {})
            sg_style = supergroups_style.get(sg_level_str, {})
            default_sg_label_size = 8 if str(sg_level_str) in ("1", "2") else int(cfg.enclosure_label_text_size)
            default_sg_label_distance = -0.05
            sg_label_font = sg_style.get("label_font", "Arial")
            sg_label_size = int(sg_style.get("label_size", default_sg_label_size))
            sg_label_color = sg_style.get("label_color", "#000000")
            sg_label_distance = _safe_float(sg_style.get("label_distance"), default_sg_label_distance)
            sg_label_rotation = int(sg_style.get("label_rotation", 0))

            # Supergroup labels use yanchor="bottom": y is the label bottom edge.
            y_anchor = new_bb_y1 + sg_label_distance * float(cfg.block_height)

            try:
                fig.layout.annotations[aidx].visible = True
                fig.layout.annotations[aidx].x = xmid
                fig.layout.annotations[aidx].y = y_anchor
                fig.layout.annotations[aidx].xanchor = "center"
                fig.layout.annotations[aidx].yanchor = "bottom"
                fig.layout.annotations[aidx].textangle = -sg_label_rotation  # Plotly uses opposite sign convention
                fig.layout.annotations[aidx].font = dict(
                    family=sg_label_font,
                    size=sg_label_size,
                    color=sg_label_color
                )
            except Exception:
                pass

            # Mask sizing using scaleanchor-aware conversion
            import math as _math_sg_mask
            _layout_sg = fig_json.get("layout", {}) if isinstance(fig_json, dict) else {}
            _xr_sg = (_layout_sg.get("xaxis") or {}).get("range", [-10, 10])
            _yr_sg = (_layout_sg.get("yaxis") or {}).get("range", [-10, 10])
            _x_span_sg = abs(float(_xr_sg[1]) - float(_xr_sg[0]))
            _y_span_sg = abs(float(_yr_sg[1]) - float(_yr_sg[0]))
            _ppu_sg = min(float(_CANVAS_W) * 0.88 / max(_x_span_sg, 0.01),
                          float(_CANVAS_H) * 0.88 / max(_y_span_sg, 0.01))   # full effective px/unit

            # --- Enclosure name truncation (max 65% of the enclosure top edge) ---
            # The cap includes the ellipsis itself (i.e., the displayed string including '…' fits within 65%).
            try:
                _top_len_data = float(new_bb_x1) - float(new_bb_x0)
            except Exception:
                _top_len_data = float(bb[1]) - float(bb[0])
            _max_w_data = max(0.0, float(_top_len_data) * 0.58)  # 58% — tighter than 0.65
            _max_w_px = float(_max_w_data) * float(_ppu_sg)

            # Estimate text box in px (single-line); rotation handled in x_extent formula below.
            angle_rad_sg = _math_sg_mask.radians(sg_label_rotation)
            text_h_px_sg = sg_label_size * 1.2
            _sin = abs(_math_sg_mask.sin(angle_rad_sg))
            _cos = abs(_math_sg_mask.cos(angle_rad_sg))
            _char_w_px = float(sg_label_size) * 0.72  # conservative: wide chars A,W,M ≈ 0.70-0.85

            txt_raw = str(txt)
            txt_use = txt_raw

            if _max_w_px > 0.0:
                # If rotation makes x-extent depend mostly on height (cos≈0), truncation is unnecessary.
                if _cos > 1e-6:
                    _budget_for_tw = float(_max_w_px) - abs(text_h_px_sg * _sin)
                    if _budget_for_tw < 0:
                        _budget_for_tw = 0.0
                    _max_tw = _budget_for_tw / float(_cos)
                    # Subtract 1-char safety buffer for font/browser render variance
                    _max_chars = max(1, int(_math_sg_mask.floor(_max_tw / max(_char_w_px, 1e-9))) - 1)
                    if len(txt_raw) > _max_chars:
                        if _max_chars <= 1:
                            txt_use = "…"
                        else:
                            _keep = max(1, int(_max_chars) - 1)
                            txt_use = txt_raw[:_keep].rstrip() + "…"

            # Apply truncated text to the annotation so the displayed text matches the cap.
            try:
                fig.layout.annotations[aidx].text = txt_use
            except Exception:
                pass

            nchars_sg = max(1, len(txt_use))
            text_w_px_sg = nchars_sg * sg_label_size * 0.72  # same ratio as _char_w_px
            x_extent_px_sg = abs(text_w_px_sg * _cos) + abs(text_h_px_sg * _sin)
            y_extent_px_sg = abs(text_w_px_sg * _sin) + abs(text_h_px_sg * _cos)
            pad_px_sg = 1.0
            half_w_sg = (x_extent_px_sg / 2.0 + pad_px_sg) / max(_ppu_sg, 1.0)
            h_data_sg = (y_extent_px_sg + pad_px_sg) / max(_ppu_sg, 1.0)

            mx0 = max(new_bb_x0, xmid - half_w_sg)
            mx1 = min(new_bb_x1, xmid + half_w_sg)
            if mx1 <= mx0:
                mx0, mx1 = xmid - 1e-6, xmid + 1e-6

            my0 = y_anchor - pad_px_sg / max(_ppu_sg, 1.0)
            my1 = y_anchor + h_data_sg

            try:
                fig.layout.shapes[midx].x0 = mx0
                fig.layout.shapes[midx].x1 = mx1
                fig.layout.shapes[midx].y0 = my0
                fig.layout.shapes[midx].y1 = my1
                fig.layout.shapes[midx].fillcolor = "#FFFFFF"
                fig.layout.shapes[midx].line = dict(color="rgba(0,0,0,0)", width=0)
                fig.layout.shapes[midx].opacity = 1.0
                fig.layout.shapes[midx].visible = True
                fig.layout.shapes[midx].layer = "below"
            except Exception:
                pass

    # === Expand viewport to accommodate shifted content ===
    # First, update meta's slice_label_base_x so apply_slice_label_style uses
    # the shifted enclosure right edge (not the original one).
    if show_group and global_shift > 0 and group_shift_map:
        orig_enc_x1 = float(meta.get("enclosure_global_x1", 0.0))
        # For each slice, the enclosure right shifts by last_shift
        # Use the maximum across all slices for consistent alignment
        max_right_shift = max(
            info.get("last_shift", 0.0) for info in group_shift_map.values()
        )
        new_enc_x1 = orig_enc_x1 + max_right_shift
        # Update per-slice base_x to use the shifted enclosure right edge
        base_x_map = meta.get("slice_label_base_x", {})
        for sid_key in base_x_map:
            base_x_map[sid_key] = new_enc_x1
        meta["enclosure_global_x1"] = new_enc_x1

    try:
        xr = list(fig.layout.xaxis.range) if hasattr(fig.layout.xaxis, "range") and fig.layout.xaxis.range else None
        yr = list(fig.layout.yaxis.range) if hasattr(fig.layout.yaxis, "range") and fig.layout.yaxis.range else None

        min_x, max_x = float("inf"), float("-inf")
        max_y = float("-inf")

        # Check all rect shapes
        for shp in fig.layout.shapes:
            if getattr(shp, "type", "") not in ("rect", "path"):
                continue
            vis = getattr(shp, "visible", True)
            if vis is False:
                continue
            rx0 = getattr(shp, "x0", None)
            rx1 = getattr(shp, "x1", None)
            ry1 = getattr(shp, "y1", None)
            if (rx0 is None) or (rx1 is None) or (ry1 is None):
                continue
            x0, x1 = _safe_float(rx0, 0.0), _safe_float(rx1, 0.0)
            y1 = _safe_float(ry1, 0.0)
            min_x, max_x = min(min_x, x0, x1), max(max_x, x0, x1)
            max_y = max(max_y, y1)

        # Check visible annotations
        for ann in (fig.layout.annotations or []):
            if not getattr(ann, "visible", False):
                continue
            x_val = float(getattr(ann, "x", 0))
            txt = str(getattr(ann, "text", "") or "")
            text_w = estimate_label_width_data_units(txt, cfg) * cfg.enclosure_group_label_width_fudge
            xanchor = str(getattr(ann, "xanchor", "center"))
            if xanchor == "right":
                min_x = min(min_x, x_val - text_w)
            elif xanchor == "left":
                max_x = max(max_x, x_val + text_w)
            else:
                min_x = min(min_x, x_val - text_w / 2)
                max_x = max(max_x, x_val + text_w / 2)

        # Check all traces for x extent
        for tr in fig.data:
            try:
                if tr.x is not None and len(tr.x) > 0:
                    tr_min_x = min(tr.x)
                    tr_max_x = max(tr.x)
                    min_x = min(min_x, tr_min_x)
                    max_x = max(max_x, tr_max_x)
            except Exception:
                pass

        if min_x != float("inf"):
            new_x_left = min_x - cfg.viewport_pad_x
            new_x_right = max_x + cfg.viewport_pad_x
            cur_x_left = xr[0] if xr else new_x_left
            cur_x_right = xr[1] if xr else new_x_right
            fig.update_xaxes(range=[min(cur_x_left, new_x_left), max(cur_x_right, new_x_right)])
        if yr and max_y != float("-inf"):
            need_y1 = max_y + cfg.viewport_pad_y
            if need_y1 > yr[1]:
                fig.update_yaxes(range=[yr[0], need_y1])

    except Exception:
        pass

    return fig.to_dict()


def _apply_truncation_names(fig_dict: dict, meta: dict) -> None:
    """
    For every enclosure-label annotation whose currently displayed text
    differs from the original (truncated/wrapped), set:

      annotation["hovertext"]   = full original text
      annotation["hoverlabel"]  = styling matching the block tooltip

    Why this combination:

      • `hovertext` is the plotly-standard channel that survives schema
        validation, `go.Figure(...).to_dict()` round-trips, and dash's
        figure serialization with no field loss. It carries the full
        text from Python all the way to gd._fullLayout.annotations[i]
        on the JS side, where the custom mousemove handler reads it.

      • `hoverlabel` mirrors the layout-level hoverlabel used for trace
        (block) tooltips so font / colour are identical between block
        and aggregation-name tooltips, in case the plotly native
        annotation tooltip ever leaks through.

    Note we do NOT set `captureevents = False`. Earlier attempts did,
    aiming to suppress plotly's native annotation tooltip rendering —
    but that flag also caused plotly.js to drop the `hovertext`
    payload from `_fullLayout.annotations`, which broke the custom
    JS handler that reads it. Plotly's native tooltip is suppressed
    instead by a CSS rule that hides the `.annotation-hovertext`
    SVG group; see app.index_string.

    Why a helper rather than inlining: the initial figure (built once at
    app startup, line ~6863) needs the fields set so the tooltip works
    on first paint, BEFORE any callback fires. render_with_highlight
    re-runs the same logic during callbacks. Both must be identical.
    Mutates fig_dict in place; tolerant of malformed input.
    """
    try:
        _by_idx = (meta or {}).get("truncated_label_by_anno_idx", {}) or {}
        _annos = ((fig_dict or {}).get("layout", {}) or {}).get("annotations", []) or []

        # Build {anno_idx_str: full_text} for ALL enclosure name labels
        # (not just truncated ones). Even labels that are currently
        # showing in full benefit from a tooltip — the user gets a
        # consistent affordance across every label, and labels that
        # appear "complete" can still be useful to confirm via tooltip
        # (e.g., to read the full text without squinting at small fonts,
        # or to verify the displayed text matches the canonical name).
        # The hit zones are computed identically; only this map decides
        # which annotations are eligible for the tooltip.
        #
        # We DO NOT set annotation.hovertext — see explanatory comment
        # at the top of this function for why we route via layout.meta.
        _enclosure_name_full_text: dict = {}
        for _aidx_str, _orig_full in _by_idx.items():
            try:
                _ai = int(_aidx_str)
                if not (0 <= _ai < len(_annos)):
                    continue
                _ann = _annos[_ai]
                if _ann is None:
                    continue
                if _orig_full and str(_orig_full).strip():
                    _enclosure_name_full_text[str(_ai)] = str(_orig_full)
                # Defensive cleanup: zero out any stale hovertext/hoverlabel
                # left over from earlier code paths. Idempotent across renders.
                if "hovertext" in _ann:
                    _ann["hovertext"] = ""
                if "hoverlabel" in _ann:
                    try:
                        del _ann["hoverlabel"]
                    except Exception:
                        pass
            except Exception:
                pass

        # Inject the {idx: full_text} map into fig.layout.meta WITHOUT
        # clobbering existing keys (base_sizes etc.). meta has valType="any",
        # so it survives schema validation, dash serialization, and plotly
        # redraws unchanged. JS reads gd._fullLayout.meta.truncated_full_text.
        # NOTE: the meta key is still named `truncated_full_text` for
        # backward compatibility with the JS reader, but it now contains
        # the full text for ALL enclosure name labels (truncated or not).
        if not isinstance(fig_dict, dict):
            return
        _layout = fig_dict.setdefault("layout", {})
        if not isinstance(_layout, dict):
            return
        _existing_meta = _layout.get("meta")
        if not isinstance(_existing_meta, dict):
            _existing_meta = {}
        _existing_meta["truncated_full_text"] = _enclosure_name_full_text
        _layout["meta"] = _existing_meta
    except Exception:
        pass


def apply_enclosure_border_visibility(fig_json: dict, meta: dict, enclosure_visibility: dict) -> dict:
    """
    Control visibility of enclosure (supergroup) borders based on enclosure_visibility state.
    enclosure_visibility is a dict like {"1": True, "2": False, ...} where keys are level numbers.

    NOTE: This function ONLY controls the border shape visibility.
    The annotation (name) and mask visibility are controlled separately by apply_enclosure_label_visibility.
    """
    if not fig_json:
        return fig_json

    enclosure_visibility = enclosure_visibility or {}

    level_keys = list(meta.get("enclosure_label_level_keys", []))
    border_indices = list(meta.get("enclosure_label_border_shape_indices", []))

    fig = go.Figure(fig_json)

    if (not hasattr(fig.layout, "shapes")) or (not fig.layout.shapes):
        return fig.to_dict()

    n = min(len(level_keys), len(border_indices))
    if n <= 0:
        return fig.to_dict()

    for i in range(n):
        k = level_keys[i]
        bidx = border_indices[i]

        # Only handle supergroup levels (not "group")
        if not k.startswith("supergroup"):
            continue

        # Extract level number from "supergroupN"
        try:
            level_str = k.replace("supergroup", "")
            level_num = int(level_str)
        except Exception:
            continue

        # Check visibility state (default True if not set)
        is_visible = enclosure_visibility.get(str(level_num), True)

        # Hide/show ONLY the border shape
        if bidx >= 0 and bidx < len(fig.layout.shapes):
            try:
                fig.layout.shapes[bidx].visible = is_visible
            except Exception:
                pass

    return fig.to_dict()


def apply_slice_label_style(fig_json: dict, meta: dict, slice_style: dict, rename_map: dict = None) -> dict:
    # Apply slice label annotation style (visibility/font/size/rotation) and optional renames.
    # rename_map: {original_sid: "display_name", ...} for user-edited slice labels.
    # IMPORTANT:
    # - Do NOT rely on annotation indices (they can change when other callbacks insert/reorder labels).
    # - Identify slice labels by a stable marker (templateitemname="SLICE_LABEL::<sid>") created at build time.
    # - Guard against accidental coupling: slice labels are identified by the marker first, then (fallback) by matching text
    #   near the stored base position.
    #   We do NOT assume slice labels always stay on the left because the "Label distance" control can move them across x=0.
    if not fig_json:
        return fig_json

    rename_map = rename_map or {}
    slice_style = slice_style or {}
    visible = bool(slice_style.get("visible", True))
    font_family = str(slice_style.get("font") or "Arial")
    try:
        font_size = int(float(slice_style.get("size", 11)))
    except Exception:
        font_size = 11
    try:
        rot = int(float(slice_style.get("rotation", 0)))
    except Exception:
        rot = 0

    color = _normalize_hex_color(str(slice_style.get("color", "#000000")), "#000000")
    try:
        dist = float(slice_style.get("distance", 0.0))
    except Exception:
        dist = 0.0
    dist = float(max(0.0, min(5.0, dist)))

    annos = (fig_json.get("layout", {}) or {}).get("annotations", []) or []

    base_x_map = dict((meta or {}).get("slice_label_base_x", {}) or {})
    base_y_map = dict((meta or {}).get("slice_label_base_y", {}) or {})

    # Pass 1: update by stable templateitemname marker (and x < 0 guard)
    updated_indices = set()
    for i, a0 in enumerate(annos):
        a = dict(a0 or {})
        tname = str(a.get("templateitemname") or "")
        if not tname.startswith("SLICE_LABEL::"):
            continue

        # Guard: slice labels are on the LEFT; if marker is on a right-side label, strip it.
        try:
            ax = float(a.get("x", 0.0))
        except Exception:
            ax = 0.0

        sid = tname.split("::", 1)[1]

        a["visible"] = visible
        a["font"] = dict(a.get("font") or {})
        a["font"]["family"] = font_family
        a["font"]["size"] = float(font_size)
        a["font"]["color"] = color
        a["textangle"] = -rot

        # Apply rename if present
        if str(sid) in rename_map:
            a["text"] = str(rename_map[str(sid)])

        if str(sid) in base_x_map:
            try:
                a["x"] = float(base_x_map.get(str(sid))) + float(dist)
            except Exception:
                pass

        annos[i] = a
        updated_indices.add(i)

    # Pass 2: (backward compatibility) if marker is missing everywhere, infer slice labels from meta.
    # Only consider LEFT-side candidates (x < 0) to avoid matching supergroup/group numeric labels.
    if meta and len(updated_indices) == 0 and base_x_map and base_y_map:
        tol = 1e-3
        for sid in base_x_map.keys():
            sid_str = str(sid)
            if sid_str not in base_y_map:
                continue
            try:
                bx = float(base_x_map.get(sid_str))
                by = float(base_y_map.get(sid_str))
            except Exception:
                continue

            best_i = None
            best_err = None

            for j, a0 in enumerate(annos):
                a = dict(a0 or {})

                # Must match the slice id text exactly.
                if str(a.get("text", "")) != sid_str:
                    continue

                try:
                    ax = float(a.get("x", 0.0))
                    ay = float(a.get("y", 0.0))
                except Exception:
                    continue

                # Must be on the LEFT side.

                # Match by proximity to base (x,y). Use current x (which may already include prior dist).
                dx = abs(ax - bx)
                dy = abs(ay - by)
                err = dx + dy

                if best_err is None or err < best_err:
                    best_err = err
                    best_i = j

            if best_i is None:
                continue

            if best_err is not None and best_err > 0.5:
                # Too far: likely not the slice label.
                continue

            a = dict(annos[best_i] or {})
            a["templateitemname"] = f"SLICE_LABEL::{sid_str}"
            a["visible"] = visible
            a["font"] = dict(a.get("font") or {})
            a["font"]["family"] = font_family
            a["font"]["size"] = float(font_size)
            a["font"]["color"] = color
            a["textangle"] = -rot
            # Apply rename if present
            if sid_str in rename_map:
                a["text"] = str(rename_map[sid_str])
            try:
                a["x"] = float(bx) + float(dist)
            except Exception:
                pass

            annos[best_i] = a
            updated_indices.add(best_i)

    fig_json.setdefault("layout", {})["annotations"] = annos

    return fig_json


def apply_btn_nonces(fig_json: dict, meta: dict, nonce_map: Dict[str, int]) -> dict:
    """
    Ensure every BTNEXO trace has a changing customdata so repeated clicks always trigger Dash.
    """
    curve_to_id = meta.get("curve_to_id", [])
    fig = go.Figure(fig_json)

    for idx, tr in enumerate(fig.data):
        if idx >= len(curve_to_id):
            continue
        oid = curve_to_id[idx]
        if not oid.startswith("BTNEXO::"):
            continue

        sid = oid.split("::", 1)[1]
        v = int(nonce_map.get(sid, 0))

        # customdata length must match number of points in the trace
        try:
            npts = len(tr.x) if tr.x is not None else 1
        except Exception:
            npts = 1

        tr.customdata = [v] * max(1, npts)

    return fig.to_dict()


def apply_click_nonces(fig_json: dict, meta: dict, nonce_map: Dict[str, int]) -> dict:
    """
    Ensure repeated clicks on the SAME block/band/background still trigger Dash callbacks,
    by injecting a changing customdata into those traces.
    """
    curve_to_id = meta.get("curve_to_id", [])
    fig = go.Figure(fig_json)

    for idx, tr in enumerate(fig.data):
        if idx >= len(curve_to_id):
            continue

        oid = curve_to_id[idx]

        # Exo buttons already handled by apply_btn_nonces
        if oid.startswith("BTNEXO::"):
            continue

        base_oid = oid.replace("::TEXT", "") if oid.endswith("::TEXT") else oid

        if (base_oid == "BG::CLICK") or base_oid.startswith("BLOCK::") or base_oid.startswith("BAND::"):
            v = int(nonce_map.get(base_oid, 0))
            try:
                npts = len(tr.x) if tr.x is not None else 1
            except Exception:
                npts = 1
            tr.customdata = [v] * max(1, npts)

    return fig.to_dict()


# 8) Dash app

def make_app(slice_paths: List[str], slice_names: Optional[List[str]], element_col: str,
             default_category_col: Optional[str]) -> Dash:
    cfg = StyleConfig()

    def _get_default_supergroup_style(level: int) -> Dict[str, Any]:
        """Get default style for a supergroup level."""
        return {
            "pad_left": (0.5 if level == 1 else 0.2),
            "pad_right": (0.5 if level == 1 else 0.2),
            "pad_top": (0.5 if level == 1 else 0.5),
            "pad_bottom": (0.5 if level == 1 else 0.5),
            "enclosure_gap": (2.0 if level == 1 else 3.0),
            "fill_color": "#FFFFFF",
            "fill_opacity": 0.0,
            "border_color": "#000000",
            "border_opacity": (0.5 if level == 1 else 0.4),
            "border_width": (1.0 if level == 1 else 0.7),
            "radii": [0, 0, 0, 0],
            "line_style": "solid",
            "label_font": "Arial",
            "label_size": 8,
            "label_color": "#000000",
            "label_distance": -0.05,
            "label_rotation": 0,
        }

    slice_paths = slice_paths or []

    # Allow starting the app with zero CLI slices (UI upload will provide data later).
    # To keep downstream logic stable (dfs_raw[0], category candidates, etc.),
    # create one placeholder slice with a minimal 2-column schema.
    if len(slice_paths) < 1:
        slice_paths = ["(UI upload)"]
        dfs_raw: List[pd.DataFrame] = [pd.DataFrame(columns=["Category", "Element"])]
    else:
        # Load all slices once
        dfs_raw = [read_slice_excel(p) for p in slice_paths]

    cols0 = list(dfs_raw[0].columns)

    # Element column is always the rightmost column, regardless of its name.
    element_col = cols0[-1]

    # Ensure consistent columns across slices.
    # If column order differs, auto-reorder to match the first slice (cols0) to keep right-adjacent logic stable.
    cols0_set = set(cols0)

    for idx_df, (p, df) in enumerate(zip(slice_paths, dfs_raw)):
        if len(df.columns) < 1:
            raise ValueError(f"Slice has no columns: {p}")

        cols_i = list(df.columns)
        cols_i_set = set(cols_i)

        if cols_i_set != cols0_set:
            missing = [c for c in cols0 if c not in cols_i_set]
            extra = [c for c in cols_i if c not in cols0_set]
            raise ValueError(
                f"Column set mismatch in slice: {p}\n"
                f"Missing columns: {missing}\n"
                f"Extra columns:   {extra}\n"
                f"Expected columns (set): {sorted(list(cols0_set))}\n"
                f"Got columns (set):      {sorted(list(cols_i_set))}"
            )

        if cols_i != cols0:
            dfs_raw[idx_df] = df[cols0].copy()

    common_cols = set(dfs_raw[0].columns)
    for df in dfs_raw[1:]:
        common_cols &= set(df.columns)

    # Category must NOT be the last column, because we use the right-adjacent column as element identity.
    category_candidates = [c for c in cols0[:-1] if (c != element_col and c in common_cols)]
    if not category_candidates:
        raise ValueError("No common category candidates found across slices.")

    category_col = default_category_col if (default_category_col in category_candidates) else category_candidates[-1]

    # Slice ids
    if slice_names and len(slice_names) == len(slice_paths):
        slice_ids = slice_names
    else:
        slice_ids = [Path(p).stem for p in slice_paths]

    def build_base(category_col_value: str, x_idx: int, y_idx: int, style_state: Optional[dict] = None,
                   upload_state: Optional[dict] = None, band_mode: str = "existence", band_proportion: float = 0.5,
                   block_mode: str = "existence", block_median_width: float = None, collapse_state: dict = None,
                   sweep_k_max: int = 10, sweep_m: int = 2, sweep_delta: float = 0.01,
                   graph_size: Optional[dict] = None, initial_y_range: Optional[list] = None):
        # IMPORTANT: rebuild should recompute Exo anchor for the new layout
        # (we apply style overrides via a per-rebuild cfg_work)

        band_state = ((style_state or {}).get("band", {}) or {})
        band_colors_in = band_state.get("colors")
        if isinstance(band_colors_in, dict) and len(band_colors_in) > 0:
            band_colors_use = {str(k): str(v) for k, v in band_colors_in.items()}
        else:
            band_colors_use = dict(cfg.band_colors)

        # Per-type band opacity / width (backward compatible with legacy scalar fields)
        try:
            band_opacity_default_use = float(band_state.get("opacity", cfg.band_default_opacity))
        except Exception:
            band_opacity_default_use = float(cfg.band_default_opacity)
        band_opacity_default_use = float(max(0.0, min(1.0, band_opacity_default_use)))

        op_in = band_state.get("opacities")
        if isinstance(op_in, dict) and len(op_in) > 0:
            band_op_map = {str(k): float(v) for k, v in op_in.items()}
        else:
            band_op_map = {str(k): float(band_opacity_default_use) for k in band_colors_use.keys()}

        for k in band_colors_use.keys():
            kk = str(k)
            try:
                vv = float(band_op_map.get(kk, float(band_opacity_default_use)))
            except Exception:
                vv = float(band_opacity_default_use)
            band_op_map[kk] = float(max(0.0, min(1.0, vv)))

        try:
            band_width_ratio_default_use = float(band_state.get("width_ratio", cfg.band_width_ratio))
        except Exception:
            band_width_ratio_default_use = float(cfg.band_width_ratio)
        band_width_ratio_default_use = float(max(0.02, min(0.30, band_width_ratio_default_use)))

        w_in = band_state.get("width_ratios")
        if isinstance(w_in, dict) and len(w_in) > 0:
            band_w_map = {str(k): float(v) for k, v in w_in.items()}
        else:
            band_w_map = {str(k): float(band_width_ratio_default_use) for k in band_colors_use.keys()}

        for k in band_colors_use.keys():
            kk = str(k)
            try:
                vv = float(band_w_map.get(kk, float(band_width_ratio_default_use)))
            except Exception:
                vv = float(band_width_ratio_default_use)
            band_w_map[kk] = float(max(0.02, min(0.30, vv)))

        layout_state = ((style_state or {}).get("layout", {}) or {})
        try:
            layer_gap_use = float(layout_state.get("layer_gap", cfg.layer_gap))
        except Exception:
            layer_gap_use = float(cfg.layer_gap)
        layer_gap_use = float(max(1.0, min(20.0, layer_gap_use)))

        try:
            exo_gap_use = float(layout_state.get("exo_gap", cfg.exo_gap))
        except Exception:
            exo_gap_use = float(cfg.exo_gap)
        exo_gap_use = float(max(0.5, min(5.0, exo_gap_use)))

        try:
            inner_gap_use = float(layout_state.get("inner_gap", cfg.inner_gap))
        except Exception:
            inner_gap_use = float(cfg.inner_gap)
        inner_gap_use = float(max(0.0, min(1.0, inner_gap_use)))

        try:
            outer_gap_use = float(layout_state.get("outer_gap", cfg.outer_gap))
        except Exception:
            outer_gap_use = float(cfg.outer_gap)
        outer_gap_use = float(max(0.0, min(10.0, outer_gap_use)))

        # ========== Process Block Style State ==========
        block_state = ((style_state or {}).get("block", {}) or {})

        # Block widths
        block_widths_map = {}
        widths_in = block_state.get("widths")
        if isinstance(widths_in, dict) and len(widths_in) > 0:
            for k, v in widths_in.items():
                try:
                    block_widths_map[str(k)] = float(max(0.5, min(5.0, float(v))))
                except Exception:
                    block_widths_map[str(k)] = float(cfg.block_width)
        for k in cfg.block_colors.keys():
            block_widths_map.setdefault(str(k), float(cfg.block_width))

        # Block heights
        block_heights_map = {}
        heights_in = block_state.get("heights")
        if isinstance(heights_in, dict) and len(heights_in) > 0:
            for k, v in heights_in.items():
                try:
                    block_heights_map[str(k)] = float(max(0.2, min(20.0, float(v))))
                except Exception:
                    block_heights_map[str(k)] = float(cfg.block_height)
        for k in cfg.block_colors.keys():
            block_heights_map.setdefault(str(k), float(cfg.block_height))

        # Block fill colors
        block_fill_colors_map = {}
        fill_colors_in = block_state.get("fill_colors")
        if isinstance(fill_colors_in, dict) and len(fill_colors_in) > 0:
            for k, v in fill_colors_in.items():
                block_fill_colors_map[str(k)] = str(v) if v else cfg.block_colors.get(str(k), "#CCCCCC")
        for k in cfg.block_colors.keys():
            block_fill_colors_map.setdefault(str(k), cfg.block_colors.get(str(k), "#CCCCCC"))

        # Block fill opacities
        block_fill_opacities_map = {}
        fill_op_in = block_state.get("fill_opacities")
        if isinstance(fill_op_in, dict) and len(fill_op_in) > 0:
            for k, v in fill_op_in.items():
                try:
                    block_fill_opacities_map[str(k)] = float(max(0.0, min(1.0, float(v))))
                except Exception:
                    block_fill_opacities_map[str(k)] = 1.0
        for k in cfg.block_colors.keys():
            block_fill_opacities_map.setdefault(str(k), 1.0)

        # Block border colors
        block_border_colors_map = {}
        border_colors_in = block_state.get("border_colors")
        if isinstance(border_colors_in, dict) and len(border_colors_in) > 0:
            for k, v in border_colors_in.items():
                block_border_colors_map[str(k)] = str(v) if v else "#222222"
        for k in cfg.block_colors.keys():
            block_border_colors_map.setdefault(str(k), "#222222")

        # Block border opacities
        block_border_opacities_map = {}
        border_op_in = block_state.get("border_opacities")
        if isinstance(border_op_in, dict) and len(border_op_in) > 0:
            for k, v in border_op_in.items():
                try:
                    block_border_opacities_map[str(k)] = float(max(0.0, min(1.0, float(v))))
                except Exception:
                    block_border_opacities_map[str(k)] = 1.0
        for k in cfg.block_colors.keys():
            block_border_opacities_map.setdefault(str(k), 1.0)

        # Block border radii
        block_border_radii_map = {}
        radii_in = block_state.get("border_radii")
        if isinstance(radii_in, dict) and len(radii_in) > 0:
            for k, v in radii_in.items():
                if isinstance(v, list) and len(v) == 4:
                    try:
                        # Handle None values individually
                        def safe_radius(r):
                            if r is None:
                                return 0
                            try:
                                return int(max(0, min(50, int(r))))
                            except (ValueError, TypeError):
                                return 0

                        block_border_radii_map[str(k)] = [safe_radius(r) for r in v]
                    except Exception:
                        block_border_radii_map[str(k)] = [0, 0, 0, 0]
                else:
                    block_border_radii_map[str(k)] = [0, 0, 0, 0]
        for k in cfg.block_colors.keys():
            block_border_radii_map.setdefault(str(k), [0, 0, 0, 0])

        # Block line styles
        block_line_styles_map = {}
        line_styles_in = block_state.get("line_styles")
        if isinstance(line_styles_in, dict) and len(line_styles_in) > 0:
            for k, v in line_styles_in.items():
                block_line_styles_map[str(k)] = str(v) if v in ("solid", "dash", "dot", "dashdot") else "solid"
        for k in cfg.block_colors.keys():
            block_line_styles_map.setdefault(str(k), "solid")

        # Block text fonts
        block_text_fonts_map = {}
        text_fonts_in = block_state.get("text_fonts")
        if isinstance(text_fonts_in, dict) and len(text_fonts_in) > 0:
            for k, v in text_fonts_in.items():
                block_text_fonts_map[str(k)] = str(v) if v else "Arial"
        for k in cfg.block_colors.keys():
            block_text_fonts_map.setdefault(str(k), "Arial")

        # Block text sizes
        block_text_sizes_map = {}
        text_sizes_in = block_state.get("text_sizes")
        if isinstance(text_sizes_in, dict) and len(text_sizes_in) > 0:
            for k, v in text_sizes_in.items():
                try:
                    block_text_sizes_map[str(k)] = int(max(8, min(72, int(v))))
                except Exception:
                    block_text_sizes_map[str(k)] = int(cfg.block_text_size)
        for k in cfg.block_colors.keys():
            block_text_sizes_map.setdefault(str(k), int(cfg.block_text_size))

        # Block text colors
        block_text_colors_map = {}
        text_colors_in = block_state.get("text_colors")
        if isinstance(text_colors_in, dict) and len(text_colors_in) > 0:
            for k, v in text_colors_in.items():
                block_text_colors_map[str(k)] = str(v) if v else cfg.block_text_colors.get(str(k), "#111111")
        for k in cfg.block_colors.keys():
            block_text_colors_map.setdefault(str(k), cfg.block_text_colors.get(str(k), "#111111"))

        # Block text aligns
        block_text_aligns_map = {}
        text_aligns_in = block_state.get("text_aligns")
        if isinstance(text_aligns_in, dict) and len(text_aligns_in) > 0:
            for k, v in text_aligns_in.items():
                block_text_aligns_map[str(k)] = str(v) if v in ("left", "center", "right") else "center"
        for k in cfg.block_colors.keys():
            block_text_aligns_map.setdefault(str(k), "center")

        # Block text rotations
        block_text_rotations_map = {}
        text_rotations_in = block_state.get("text_rotations")
        if isinstance(text_rotations_in, dict) and len(text_rotations_in) > 0:
            for k, v in text_rotations_in.items():
                try:
                    rot = int(v)
                    rot = max(-180, min(180, rot))
                    block_text_rotations_map[str(k)] = rot
                except Exception:
                    block_text_rotations_map[str(k)] = 90
        for k in cfg.block_colors.keys():
            block_text_rotations_map.setdefault(str(k), 90)

        # Block line spacings
        block_line_spacings_map = {}
        line_spacings_in = block_state.get("line_spacings")
        if isinstance(line_spacings_in, dict) and len(line_spacings_in) > 0:
            for k, v in line_spacings_in.items():
                try:
                    ls = float(v)
                    ls = max(-20, min(200, ls))
                    block_line_spacings_map[str(k)] = ls
                except Exception:
                    block_line_spacings_map[str(k)] = 0
        for k in cfg.block_colors.keys():
            block_line_spacings_map.setdefault(str(k), 0)

        # Block border widths
        block_border_widths_map = {}
        border_widths_in = block_state.get("border_widths")
        if isinstance(border_widths_in, dict) and len(border_widths_in) > 0:
            for k, v in border_widths_in.items():
                try:
                    block_border_widths_map[str(k)] = float(max(0.0, min(5.0, float(v))))
                except Exception:
                    block_border_widths_map[str(k)] = float(cfg.block_border_width)
        for k in cfg.block_colors.keys():
            block_border_widths_map.setdefault(str(k), float(cfg.block_border_width))

        # IMPORTANT: block width/height are global GEOMETRY controls used across layout.
        # The style panel stores per-type maps, but the layout engine primarily consults
        # cfg.block_width / cfg.block_height. Therefore, we promote a representative
        # value from the per-type maps into cfg_work.block_width / cfg_work.block_height.
        def _representative_size(m: Dict[str, float], default_v: float) -> float:
            try:
                vals = [float(v) for v in (m or {}).values() if v is not None]
            except Exception:
                vals = []
            if not vals:
                return float(default_v)
            first = float(vals[0])
            if all(abs(float(v) - first) < 1e-9 for v in vals):
                return float(first)
            vals_sorted = sorted([float(v) for v in vals])
            return float(vals_sorted[len(vals_sorted) // 2])

        block_width_use = _representative_size(block_widths_map, float(cfg.block_width))
        block_height_use = _representative_size(block_heights_map, float(cfg.block_height))

        cfg_work = dataclasses.replace(
            cfg,
            block_width=block_width_use,
            block_height=block_height_use,
            band_colors=band_colors_use,
            band_default_opacity=band_opacity_default_use,
            band_width_ratio=band_width_ratio_default_use,
            layer_gap=layer_gap_use,
            exo_gap=exo_gap_use,
            inner_gap=inner_gap_use,
            outer_gap=outer_gap_use,
        )
        # Attach enclosure gap map from supergroup styles
        supergroups_style = (style_state or {}).get("supergroups", {}) or {}
        enclosure_gaps = {}
        for lvl_str, lvl_style in supergroups_style.items():
            try:
                lvl_int = int(lvl_str)
                default_gap = 2.0 if lvl_int == 1 else 3.0
                enclosure_gaps[lvl_int] = float(lvl_style.get("enclosure_gap", default_gap))
            except Exception:
                pass
        cfg_work.enclosure_gaps = enclosure_gaps
        # Attach per-type maps (used by band rendering if supported)
        cfg_work.band_opacity_by_type = dict(band_op_map)
        cfg_work.band_width_ratio_by_type = dict(band_w_map)

        # Attach per-type block style maps
        cfg_work.block_widths_by_type = dict(block_widths_map)
        cfg_work.block_heights_by_type = dict(block_heights_map)
        cfg_work.block_fill_colors_by_type = dict(block_fill_colors_map)
        cfg_work.block_fill_opacities_by_type = dict(block_fill_opacities_map)
        cfg_work.block_border_colors_by_type = dict(block_border_colors_map)
        cfg_work.block_border_opacities_by_type = dict(block_border_opacities_map)
        cfg_work.block_border_radii_by_type = dict(block_border_radii_map)
        cfg_work.block_line_styles_by_type = dict(block_line_styles_map)
        cfg_work.block_text_fonts_by_type = dict(block_text_fonts_map)
        cfg_work.block_text_sizes_by_type = dict(block_text_sizes_map)
        cfg_work.block_text_colors_by_type = dict(block_text_colors_map)
        cfg_work.block_text_aligns_by_type = dict(block_text_aligns_map)
        cfg_work.block_text_rotations_by_type = dict(block_text_rotations_map)
        cfg_work.block_line_spacings_by_type = dict(block_line_spacings_map)
        cfg_work.block_border_widths_by_type = dict(block_border_widths_map)

        # Block text padding (global from ALL panel)
        text_pads_state = block_state.get("text_pads") or {}

        def _sf_pad(v, fallback=0.0):
            try:
                return float(max(0.0, min(5.0, float(v))))
            except Exception:
                return float(fallback)

        cfg_work.block_text_pad_l = _sf_pad(text_pads_state.get("l", 0.0))
        cfg_work.block_text_pad_r = _sf_pad(text_pads_state.get("r", 0.0))
        cfg_work.block_text_pad_t = _sf_pad(text_pads_state.get("t", 0.0))
        cfg_work.block_text_pad_b = _sf_pad(text_pads_state.get("b", 0.0))

        # ========== Process Group Style State ==========
        group_state = ((style_state or {}).get("group", {}) or {})

        # Helper function to safely parse radii list
        def _safe_radii(r):
            if not r or not isinstance(r, list) or len(r) != 4:
                return [0, 0, 0, 0]
            result = []
            for val in r:
                if val is None:
                    result.append(0)
                else:
                    try:
                        result.append(int(max(0, min(50, int(val)))))
                    except (ValueError, TypeError):
                        result.append(0)
            return result

        def _safe_float_cfg(v, default):
            try:
                if v is None or v == "" or v == "None":
                    return float(default)
                return float(v)
            except (ValueError, TypeError):
                return float(default)

        cfg_work.group_style = {
            "line_width": _safe_float_cfg(group_state.get("line_width"), cfg.group_enclosure_line_width),
            "opacity": _safe_float_cfg(group_state.get("opacity"), cfg.group_enclosure_opacity),
            "color": str(group_state.get("color") or "#000000"),
            "radii": _safe_radii(group_state.get("radii", [0, 0, 0, 0])),
            "line_style": str(group_state.get("line_style") or "solid"),
        }

        # ========== Process Supergroup Styles State ==========
        supergroups_state = ((style_state or {}).get("supergroups", {}) or {})
        cfg_work.supergroup_styles = {}
        for level_str, sg_style in supergroups_state.items():
            try:
                level_int = int(level_str)
            except Exception:
                continue
            d0 = _get_default_supergroup_style(int(level_int))
            cfg_work.supergroup_styles[level_int] = {
                "pad_left": _safe_float_cfg(sg_style.get("pad_left"), d0.get("pad_left", 0.0)),
                "pad_right": _safe_float_cfg(sg_style.get("pad_right"), d0.get("pad_right", 0.0)),
                "pad_top": _safe_float_cfg(sg_style.get("pad_top"), d0.get("pad_top", 0.0)),
                "pad_bottom": _safe_float_cfg(sg_style.get("pad_bottom"), d0.get("pad_bottom", 0.0)),
                "fill_color": str(sg_style.get("fill_color") or "#FFFFFF"),
                "fill_opacity": _safe_float_cfg(sg_style.get("fill_opacity"), 0.0),
                "border_color": str(sg_style.get("border_color") or "#000000"),
                "border_opacity": _safe_float_cfg(sg_style.get("border_opacity"), 0.35),
                "border_width": _safe_float_cfg(sg_style.get("border_width"), 1.0),
                "radii": _safe_radii(sg_style.get("radii", [0, 0, 0, 0])),
                "line_style": str(sg_style.get("line_style") or "solid"),
            }

        cfg_work.exo_anchor_x = None

        # If user uploaded slices, prefer them over the CLI-provided slice_paths/dfs_raw.
        slice_paths_use = slice_paths
        slice_ids_use = slice_ids
        dfs_raw_use = dfs_raw

        if upload_state and upload_state.get("paths"):
            try:
                slice_paths_use = list(upload_state.get("paths") or [])
                slice_ids_use = list(upload_state.get("slice_names") or [Path(p).stem for p in slice_paths_use])
                dfs_raw_use = [read_slice_excel(p) for p in slice_paths_use]
            except Exception as e:
                raise ValueError(f"Failed to load uploaded slice files: {e}")

        # Defensive: Dash may pass None/out-of-range temporarily during options refresh.
        max_idx = max(0, len(slice_ids_use) - 1)

        x_idx2 = int(x_idx) if (x_idx is not None) else 0
        y_idx2 = int(y_idx) if (y_idx is not None) else (1 if max_idx >= 1 else 0)

        if x_idx2 < 0:
            x_idx2 = 0
        if x_idx2 > max_idx:
            x_idx2 = max_idx

        if y_idx2 < 0:
            y_idx2 = 0
        if y_idx2 > max_idx:
            y_idx2 = max_idx

        if y_idx2 == x_idx2:
            y_idx2 = min(x_idx2 + 1, max_idx)

        # Element column is always the rightmost column of the current data.
        element_col = list(dfs_raw_use[0].columns)[-1]

        # Per-slice prep (column existence)
        _ = validate_columns_for_choice(list(dfs_raw_use[0].columns), category_col_value, element_col)

        slices: List[SliceModel] = []

        for sid, df_raw in zip(slice_ids_use, dfs_raw_use):
            cols_raw = list(df_raw.columns)
            # Element column is always the rightmost column (per-slice, in case column order differs)
            element_col_i = cols_raw[-1]

            # Cross-level shift: element identity = the column right next to the selected category
            element_id_col_i = right_adjacent_column(cols_raw, category_col_value)

            # Fill-down: exclude only the rightmost (element) column
            df_filled = fill_down_per_slice(df_raw, exclude_cols=[element_col_i])

            if category_col_value not in df_filled.columns:
                raise ValueError(f"Category '{category_col_value}' not present in slice '{sid}'.")

            cols_i = list(df_filled.columns)

            # Higher-level columns (left of category) for enclosure/grouping
            agg_cols_i = validate_columns_for_choice(cols_i, category_col_value, element_col_i)

            # Filter out constant / useless hierarchy cols (prevents phantom "all glued" case)
            agg_cols_i2: List[str] = []
            for c in agg_cols_i:
                s = df_filled[c]
                nun = s.dropna().astype(str).str.strip().nunique()
                if nun <= 1:
                    continue
                agg_cols_i2.append(c)
            agg_cols_i = agg_cols_i2

            slices.append(build_slice_model(df_filled, sid, category_col_value, element_id_col_i, agg_cols_i))

        # Ensure cfg_work.supergroup_styles includes defaults for all supergroup levels present in data,
        # so padding is correct even before any UI interaction populates store-style.
        max_sg_level = 0
        for sl in slices:
            try:
                max_sg_level = max(int(max_sg_level), max(0, int(len(sl.agg_cols)) - 1))
            except Exception:
                pass

        for lvl in range(1, int(max_sg_level) + 1):
            if int(lvl) not in cfg_work.supergroup_styles:
                d = _get_default_supergroup_style(int(lvl))
                cfg_work.supergroup_styles[int(lvl)] = {
                    "pad_left": _safe_float_cfg(d.get("pad_left"), 0.0),
                    "pad_right": _safe_float_cfg(d.get("pad_right"), 0.0),
                    "pad_top": _safe_float_cfg(d.get("pad_top"), 0.0),
                    "pad_bottom": _safe_float_cfg(d.get("pad_bottom"), 0.0),
                    "fill_color": str(d.get("fill_color") or "#FFFFFF"),
                    "fill_opacity": _safe_float_cfg(d.get("fill_opacity"), 0.0),
                    "border_color": str(d.get("border_color") or "#000000"),
                    "border_opacity": _safe_float_cfg(d.get("border_opacity"), 0.35),
                    "border_width": _safe_float_cfg(d.get("border_width"), 1.0),
                    "radii": _safe_radii(d.get("radii", [0, 0, 0, 0])),
                    "line_style": str(d.get("line_style") or "solid"),
                }
            else:
                cur = dict(cfg_work.supergroup_styles.get(int(lvl), {}) or {})
                d = _get_default_supergroup_style(int(lvl))
                if cur.get("pad_left", None) is None:
                    cur["pad_left"] = _safe_float_cfg(d.get("pad_left"), 0.0)
                if cur.get("pad_right", None) is None:
                    cur["pad_right"] = _safe_float_cfg(d.get("pad_right"), 0.0)
                if cur.get("pad_top", None) is None:
                    cur["pad_top"] = _safe_float_cfg(d.get("pad_top"), 0.0)
                if cur.get("pad_bottom", None) is None:
                    cur["pad_bottom"] = _safe_float_cfg(d.get("pad_bottom"), 0.0)
                if cur.get("fill_color", None) is None:
                    cur["fill_color"] = str(d.get("fill_color") or "#FFFFFF")
                if cur.get("fill_opacity", None) is None:
                    cur["fill_opacity"] = _safe_float_cfg(d.get("fill_opacity"), 0.0)
                if cur.get("border_color", None) is None:
                    cur["border_color"] = str(d.get("border_color") or "#000000")
                if cur.get("border_opacity", None) is None:
                    cur["border_opacity"] = _safe_float_cfg(d.get("border_opacity"), 0.35)
                if cur.get("border_width", None) is None:
                    cur["border_width"] = _safe_float_cfg(d.get("border_width"), 1.0)
                if cur.get("radii", None) is None:
                    cur["radii"] = _safe_radii(d.get("radii", [0, 0, 0, 0]))
                if cur.get("line_style", None) is None:
                    cur["line_style"] = str(d.get("line_style") or "solid")
                cfg_work.supergroup_styles[int(lvl)] = cur

        steps = compute_all_steps(slices)

        # Use collapse_state to determine collapse anchors
        if collapse_state and collapse_state.get("collapsed", False):
            _ax = collapse_state.get("anchor_x")
            _ay = collapse_state.get("anchor_y")
            if _ax is not None and _ay is not None:
                x_idx2 = int(_ax)
                y_idx2 = int(_ay)
            else:
                x_idx2 = int(x_idx)
                y_idx2 = int(y_idx)
        else:
            x_idx2 = int(x_idx)
            y_idx2 = int(y_idx)

        slices_view, steps_view, col_step = collapse_steps(slices, steps, int(x_idx2), int(y_idx2))

        # Per PDF Section V-C: collapse only regenerates bands between the two
        # anchor slices.  All other slices keep the ordering from the original
        # (non-collapsed) full sweep.  The anchor pair is then locally reordered
        # with a single pass (Eq. 6 & 7), NOT a multi-iteration sweep.
        if col_step is not None:
            # Step 1: full sweep on ORIGINAL slices (before collapse)
            base_order = sweep_reorder(slices, steps, K_max=sweep_k_max, m=sweep_m, delta=sweep_delta)
            # Step 2: single-pass reorder for anchor pair; others keep base_order
            order = collapse_reorder(slices_view, steps_view, col_step, base_order)
        else:
            order = sweep_reorder(slices_view, steps_view, K_max=sweep_k_max, m=sweep_m, delta=sweep_delta)
        block_labels, band_labels = label_blocks_and_bands(slices_view, steps_view, cfg_work)

        fig, meta = build_figure(slices_view, steps_view, order, block_labels, band_labels, cfg_work,
                                 band_mode=band_mode, band_proportion=band_proportion, block_mode=block_mode,
                                 block_median_width=block_median_width, collapse_state=collapse_state,
                                 graph_size=graph_size, initial_y_range=initial_y_range)
        meta["category_col"] = category_col_value
        meta["element_id_col"] = right_adjacent_column(list(dfs_raw_use[0].columns), category_col_value)
        meta["collapse"] = {"x": int(x_idx2), "y": int(y_idx2)}
        meta["collapse_state"] = collapse_state  # Store full collapse state
        meta["band_mode"] = band_mode  # Store band_mode in meta for reference
        meta["band_proportion"] = band_proportion  # Store band_proportion in meta
        meta["block_mode"] = block_mode  # Store block_mode in meta
        meta["block_median_width"] = block_median_width  # Store block_median_width in meta

        # Embed base_sizes into figure layout.meta so JS can do real-time zoom scaling
        # without a Python round-trip. JS reads gd._fullLayout.meta.base_sizes and
        # applies scale on every plotly_relayout event.
        fig_dict_tmp = fig.to_dict()
        _ix = (meta.get("initial_x_range") or fig_dict_tmp.get("layout", {}).get("xaxis", {}).get("range") or [0.0,
                                                                                                               1.0])
        _init_span = abs(float(_ix[1]) - float(_ix[0])) if len(_ix) == 2 else 1.0
        _lay_font = ((fig_dict_tmp.get("layout", {}) or {}).get("font") or {}).get("size")
        _hov_font = (((fig_dict_tmp.get("layout", {}) or {}).get("hoverlabel") or {}).get("font") or {}).get("size")
        base_sizes = {
            "init_x_span": _init_span,
            "layout_font_size": float(_lay_font) if _lay_font else float(cfg.ui_font_size),
            "hoverlabel_font_size": float(_hov_font) if _hov_font else float(cfg.tooltip_font_size),
            "traces": [
                {
                    "tf": (t.get("textfont") or {}).get("size") or None,
                    "lw": (t.get("line") or {}).get("width") or None,
                    "mlw": ((t.get("marker") or {}).get("line") or {}).get("width") or None,
                }
                for t in fig_dict_tmp.get("data", [])
            ],
            "shapes": [
                {"lw": (s.get("line") or {}).get("width") or None}
                for s in fig_dict_tmp.get("layout", {}).get("shapes", [])
            ],
            "annos": [
                {"fs": (a.get("font") or {}).get("size") or None}
                for a in fig_dict_tmp.get("layout", {}).get("annotations", [])
            ],
        }
        fig.update_layout(meta=base_sizes)
        return fig, meta

    fig0, meta0 = build_base(category_col, 0, 0)

    # Apply initial label visibility: Aggregation name (group) ON, Enclosure names (supergroup*) OFF
    _level_state0_init = {
        "show": {
            str(lv.get("key", "")): (str(lv.get("key", "")) == "group")
            for lv in (meta0.get("enclosure_levels", []) or [])
        }
    }
    fig0 = apply_enclosure_label_visibility(fig0.to_dict(), meta0, _level_state0_init, cfg, graph_size=None)
    # Stamp annotation.name with the original full text on every truncated/
    # wrapped label, so the JS tooltip layer can display it on hover from
    # the very first paint — before any callback fires. (The same helper
    # is invoked inside render_with_highlight for callback-time updates.)
    _apply_truncation_names(fig0, meta0)
    import plotly.graph_objects as _go_tmp
    fig0 = _go_tmp.Figure(fig0)

    app = Dash(__name__, suppress_callback_exceptions=True)
    print("RUNNING FILE: birdcage_diagram.py  (Band panel enabled)")
    app.title = "Birdcage Diagram"

    combo_btn_h = int(round(cfg.button_font_size * 1.2 + 20))
    dropdown_control_h = int(round(cfg.ui_font_size * 1.25 + 18))
    dropdown_value_line_h = int(dropdown_control_h - 2)
    dropdown_option_line_h = int(round(cfg.ui_font_size * 1.20))
    dropdown_option_pad_v = int(round(cfg.ui_font_size * 0.30))
    dropdown_option_h = 60

    # Single-file CSS injection (no extra files), plus global UI font scaling
    # language=HTML
    app.index_string = f"""
<!DOCTYPE html>
<html>
<head>
    {{%metas%}}
    <title>{{%title%}}</title>
    {{%favicon%}}
    {{%css%}}
    <style>
      /* Disable page-level scrolling */
      html, body {{
        margin: 0;
        padding: 0;
        overflow: hidden !important;
        height: 100vh;
        width: 100vw;
      }}
      body {{
        font-size: {cfg.ui_font_size}px;
      }}
      label {{
        font-size: {cfg.ui_font_size}px;
      }}
      button {{
        font-size: 40px;
        padding: 0 14px;
        height: 64px;
        background-color: #F2F2F2;
        border: 2px solid #A0A0A0;
        border-radius: 6px;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        box-sizing: border-box;
        cursor: pointer;
      }}
      button:hover {{
        background-color: #E8E8E8;
      }}
      /* React-Select (dcc.Dropdown) */
      .Select-control, .Select-menu-outer, .Select-value, .Select-value-label, .Select-option, .Select-placeholder {{
        font-size: 40px !important;
      }}
      .Select-control {{
        height: {dropdown_control_h}px !important;
        min-height: {dropdown_control_h}px !important;
      }}
      .Select-placeholder,
      .Select--single > .Select-control .Select-value,
      .Select--single > .Select-control .Select-value-label {{
        line-height: {dropdown_value_line_h}px !important;
      }}
      .Select-input {{
        height: {dropdown_control_h}px !important;
      }}
      .Select-input > input {{
        line-height: {dropdown_value_line_h}px !important;
        padding-top: 0px !important;
        padding-bottom: 0px !important;
      }}
      .Select--single > .Select-control .Select-value {{
        height: {dropdown_control_h}px !important;
        display: flex !important;
        align-items: center !important;
      }}
      .Select-menu-outer {{
        z-index: 10000 !important;
      }}
      .Select-option,
      .VirtualizedSelectOption,
      .VirtualizedSelectFocusedOption,
      .VirtualizedSelectSelectedOption {{
        height: {dropdown_option_h}px !important;
        min-height: {dropdown_option_h}px !important;
        padding-top: 0px !important;
        padding-bottom: 0px !important;
        padding-left: 12px !important;
        padding-right: 12px !important;
        line-height: {dropdown_option_line_h}px !important;
        display: flex !important;
        align-items: center !important;
      }}


      /* Category selection combo: one button-like box = (label + dropdown) */
      .category-combo-btn {{
        display: inline-flex;
        align-items: center;
        overflow: visible;
        background: #F2F2F2;
        gap: 0;
        /* Outer frame: makes the label+dropdown visually read as one unified
           control (a "pill" with a label on the left and the selected value
           on the right). The inner label/dropdown borders are light so they
           don't fight this outer frame. */
        border: 2px solid #A0A0A0;
        border-radius: 6px;
        box-shadow: 0 1px 2px rgba(0,0,0,0.04);
      }}
      .category-combo-label {{
        padding: 0px 14px;
        font-size: 40px;
        white-space: nowrap;
        height: 100%;
        display: inline-flex;
        align-items: center;
        background: #FFFFFF;
        /* No outer borders — the parent .category-combo-btn provides the frame.
           The vertical divider between the label and the dropdown is drawn via
           border-right here so we get ONE line, not a double-up of label-right
           + dropdown-left. */
        border: none;
        border-right: 2px solid #A0A0A0;
        border-radius: 6px 0 0 6px;
        box-sizing: border-box;
      }}

      .category-combo-dropdown {{
        min-width: 200px;
        width: auto;
        flex: 1 1 auto;
      }}
      .category-combo-dropdown .Select {{
        width: 100% !important;
      }}
      .category-combo-dropdown .Select-control {{
        width: 100% !important;
        /* Parent .category-combo-btn already has an outer border; removing
           the inner one avoids the heavy double-stroke that was visible on
           the dropdown side of the combo. */
        border: none !important;
        background: transparent !important;
        border-radius: 0 6px 6px 0 !important;
        box-shadow: none !important;
      }}

      #ui-top-wrap .category-combo-dropdown .Select-menu-outer {{
        z-index: 10000 !important;
        /* Let the menu grow wider than the selector if any option is longer.
           width:auto + min-width:100% means:
             - the menu is AT LEAST as wide as the input (no shrinking)
             - but can grow to fit the longest option text
           Combined with white-space:nowrap on .Select-option below, this
           stops long level names from being truncated when the dropdown is
           opened. */
        width: auto !important;
        min-width: 100% !important;
        /* Show every option without ANY scrollbar, regardless of count.
           react-select v1 imposes max-height defaults that have to be
           neutralised on both the outer wrapper and the inner .Select-menu
           layer. */
        max-height: none !important;
        overflow: visible !important;
      }}
      /* dcc.Dropdown actually uses react-virtualized-select, which has
         multiple inner layers (Grid, VirtualizedSelectGrid, ScrollContent…)
         that EACH carry their own max-height / overflow inline-style. Using
         a wildcard descendant selector is the only reliable way to override
         every layer at once — see react-select issue #704. */
      #ui-top-wrap .category-combo-dropdown .Select-menu-outer * {{
        max-height: none !important;
        overflow: visible !important;
      }}
      #ui-top-wrap .category-combo-dropdown .Select-menu {{
        max-height: none !important;
        overflow: visible !important;
      }}
      #ui-top-wrap .category-combo-dropdown .Select-option,
      #ui-top-wrap .category-combo-dropdown .VirtualizedSelectOption {{
        white-space: nowrap !important;
      }}

      /* Graph wrap - canvas-like area */
      #graph-wrap {{
        cursor: default;
        z-index: 1;
      }}
      #graph-wrap.panning {{
        cursor: grabbing !important;
      }}

      /* Main layout z-index hierarchy: graph (bottom) < top bar, right panel (top) */
      #graph-row {{
        z-index: 1;
        position: relative;
      }}
      #ui-top-wrap {{
        position: relative;
        z-index: 100;
        background-color: white !important;
      }}
      #right-panel {{
        position: relative;
        z-index: 100;
        background-color: white !important;
      }}

      /* UI scaling support */
      #ui-top-wrap, #right-panel-inner {{
        transition: transform 0.05s ease-out;
      }}

      /* Right panel overflow control */
      #right-panel {{
        overflow-x: hidden !important;
      }}
      #right-panel-inner {{
        width: 100% !important;
        max-width: 100% !important;
        box-sizing: border-box !important;
      }}
      #right-panel-inner > div {{
        width: 100% !important;
        max-width: 100% !important;
        box-sizing: border-box !important;
      }}

      /* Force arrow cursor inside Plotly graph (avoid crosshair / drag cursors) */
      #graph .cursor-crosshair,
      #graph .nsewdrag,
      #graph .drag,
      #graph .draglayer,
      #graph .zoomlayer,
      #graph .selectlayer,
      #graph .cursor-ew-resize,
      #graph .cursor-ns-resize,
      #graph .cursor-move,
      #graph .main-svg,
      #graph .svg-container,
      #graph .plotly {{
        cursor: default !important;
      }}
      #graph * {{ cursor: default !important; }}

      /* =========================
         Print-to-PDF (browser print)
         ========================= */
      @media print {{
        html, body {{
          overflow: visible !important;
          height: auto !important;
          width: auto !important;
        }}
        #ui-top-wrap {{
          display: none !important;
        }}
        #right-panel {{
          display: none !important;
        }}
        #graph-row {{
          overflow: visible !important;
          height: auto !important;
        }}
        #graph-wrap {{
          overflow: visible !important;
          position: relative !important;
          height: auto !important;
        }}
      }}

      body.print-mode {{
        overflow: visible !important;
      }}
      body.print-mode #ui-top-wrap {{
        display: none !important;
      }}
      body.print-mode #right-panel {{
        display: none !important;
      }}
      body.print-mode #graph-row {{
        overflow: visible !important;
        height: auto !important;
      }}
      body.print-mode #graph-wrap {{
        overflow: visible !important;
        position: relative !important;
        height: auto !important;
      }}

            /* Slider + numeric input styling (right panel) */
      #right-panel-inner .rc-slider {{
        padding: 0px !important;
        margin: 0px !important;
        width: 100% !important;
        box-sizing: border-box !important;
        height: 48px !important;
      }}
      #right-panel-inner .rc-slider-rail {{
        height: 5px !important;
        border-radius: 0px !important;
        background: #B0B0B0 !important;
        top: 30px !important;
      }}
      #right-panel-inner .rc-slider-track {{
        height: 5px !important;
        border-radius: 0px !important;
        background: #505050 !important;
        top: 30px !important;
      }}
      #right-panel-inner .rc-slider-handle {{
        width: 6px !important;
        height: 40px !important;
        top: 12px !important;
        margin-top: 0px !important;
        margin-left: -3px !important;
        border: none !important;
        border-radius: 0px !important;
        background: #505050 !important;
        box-shadow: none !important;
      }}
      #right-panel-inner .rc-slider-handle:hover {{
        box-shadow: none !important;
        background: #404040 !important;
        border-color: transparent !important;
      }}
      #right-panel-inner .rc-slider-handle:active {{
        box-shadow: none !important;
        background: #303030 !important;
        border-color: transparent !important;
      }}
      #right-panel-inner .rc-slider-handle:focus {{
        outline: none !important;
        box-shadow: none !important;
        border-color: transparent !important;
      }}
      /* Hide all marks/dots/steps/tooltip */
      #right-panel-inner .rc-slider-mark,
      #right-panel-inner .rc-slider-step,
      #right-panel-inner .rc-slider-dot,
      #right-panel-inner .rc-slider-dot-active,
      #right-panel-inner .rc-slider-mark-text,
      #right-panel-inner .rc-slider-tooltip,
      .rc-slider-tooltip {{
        display: none !important;
        visibility: hidden !important;
        opacity: 0 !important;
      }}
      #right-panel-inner input[type="number"] {{
        height: 48px !important;
        padding: 6px 10px !important;
        border: 2px solid #A0A0A0 !important;
        border-radius: 6px !important;
        background: #FFFFFF !important;
        box-shadow: none !important;
        font-size: 40px !important;
        -webkit-appearance: textfield !important;
        -moz-appearance: textfield !important;
        appearance: textfield !important;
      }}
      #right-panel-inner input[type="number"]:focus {{
        outline: none !important;
        border-color: #707070 !important;
      }}

      /* Hide native number spinners */
      #right-panel-inner input[type="number"]::-webkit-inner-spin-button,
      #right-panel-inner input[type="number"]::-webkit-outer-spin-button {{
        -webkit-appearance: none !important;
        margin: 0 !important;
        display: none !important;
      }}

      /* Color picker - square shape */
      #right-panel-inner input[type="color"] {{
        width: 48px !important;
        height: 48px !important;
        padding: 2px !important;
        border: 2px solid #A0A0A0 !important;
        border-radius: 6px !important;
        cursor: pointer !important;
        flex-shrink: 0 !important;
      }}

      /* Text input (including color hex text) */
      #right-panel-inner input[type="text"] {{
        height: 48px !important;
        padding: 6px 10px !important;
        border: 2px solid #A0A0A0 !important;
        border-radius: 6px !important;
        background: #FFFFFF !important;
        box-shadow: none !important;
        font-size: 40px !important;
      }}
      #right-panel-inner input[type="text"]:focus {{
        outline: none !important;
        border-color: #707070 !important;
      }}

      /* Dropdown styling */
      /* Ensure dropdowns can stretch to the right edge of the control column */
      #right-panel-inner .dash-dropdown,
      #right-panel-inner .Select {{
        width: 100% !important;
        min-width: 0 !important;
        max-width: none !important;
        flex: 1 1 auto !important;
        box-sizing: border-box !important;
      }}
      #right-panel-inner .Select-control {{
        height: 64px !important;
        min-height: 64px !important;
        box-sizing: border-box !important;
        border: 2px solid #A0A0A0 !important;
        border-radius: 6px !important;
        display: flex !important;
        align-items: center !important;
        position: relative !important;
      }}
      /* Override the global `.Select--single > .Select-control .Select-value`
         line-height (which is set to dropdown_value_line_h ≈ 25px from the
         smaller cfg.ui_font_size). The right-panel's hardcoded font-size:40px
         needs ~48px line-height to render glyphs without bottom clipping;
         using a more specific selector chain beats the global rule.

         Also override .Select-value's HEIGHT — the global rule sets it to
         dropdown_control_h (~27px) which is too short for a 48px line-box,
         so the text gets clipped by the parent's overflow:hidden. Setting
         height matching the line-height fixes the clipping. */
      #right-panel-inner.Select--single > .Select-control .Select-value,
      #right-panel-inner .Select--single > .Select-control .Select-value,
      #right-panel-inner.Select--single > .Select-control .Select-value-label,
      #right-panel-inner .Select--single > .Select-control .Select-value-label,
      #right-panel-inner .Select-value,
      #right-panel-inner .Select-placeholder {{
        padding-left: 10px !important;
        font-size: 40px !important;
        line-height: 48px !important;
        height: 48px !important;
        padding-top: 0 !important;
        padding-bottom: 0 !important;
        position: absolute !important;
        top: 50% !important;
        transform: translateY(-50%) !important;
      }}
      #right-panel-inner .Select-value-label {{
        font-size: 40px !important;
      }}
      #right-panel-inner .Select-input {{
        height: 46px !important;
        display: flex !important;
        align-items: center !important;
      }}
      #right-panel-inner .Select-input > input {{
        height: 46px !important;
        line-height: 46px !important;
        padding: 0 !important;
      }}
      /* Dropdown menu options - tighter spacing */
      #right-panel-inner .Select-option,
      #right-panel-inner .VirtualizedSelectOption,
      #right-panel-inner .VirtualizedSelectFocusedOption,
      #right-panel-inner .VirtualizedSelectSelectedOption {{
        height: 44px !important;
        min-height: 44px !important;
        font-size: 40px !important;
        line-height: 44px !important;
        padding: 0 10px !important;
        display: flex !important;
        align-items: center !important;
      }}
      /* Dropdown arrow for right panel */
      #right-panel-inner .Select-arrow-zone {{
        position: absolute !important;
        right: 0 !important;
        top: 0 !important;
        bottom: 0 !important;
        width: 40px !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
      }}
      #right-panel-inner .Select-arrow {{
        border-color: #555 transparent transparent !important;
        border-style: solid !important;
        border-width: 10px 8px 0 !important;
      }}

      /* Button styling in right panel (exclude spinner buttons) */
      #right-panel-inner button:not(.spinner-btn) {{
        height: 64px !important;
        font-size: 40px !important;
        padding: 0 14px !important;
        border: 2px solid #A0A0A0 !important;
        border-radius: 6px !important;
        background: #F2F2F2 !important;
        box-sizing: border-box !important;
        cursor: pointer !important;
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
      }}
      #right-panel-inner button:not(.spinner-btn):hover {{
        background: #E8E8E8 !important;
      }}

      /* Top bar dropdown styling - same as right panel */
      #ui-top-wrap .Select-control {{
        height: 64px !important;
        min-height: 64px !important;
        box-sizing: border-box !important;
        border: 2px solid #A0A0A0 !important;
        border-radius: 6px !important;
        display: flex !important;
        align-items: center !important;
        position: relative !important;
      }}
      #ui-top-wrap .Select-value,
      #ui-top-wrap .Select-placeholder {{
        padding-left: 10px !important;
        font-size: 40px !important;
        padding-top: 0 !important;
        padding-bottom: 0 !important;
        position: absolute !important;
        top: 50% !important;
        transform: translateY(-50%) !important;
      }}
      #ui-top-wrap .Select-value-label {{
        font-size: 40px !important;
      }}
      #ui-top-wrap .Select-input {{
        height: 62px !important;
        display: flex !important;
        align-items: center !important;
      }}
      #ui-top-wrap .Select-input > input {{
        height: 62px !important;
        line-height: 62px !important;
        padding: 0 !important;
      }}
      /* Top bar dropdown menu options - tighter spacing */
      #ui-top-wrap .Select-option,
      #ui-top-wrap .VirtualizedSelectOption,
      #ui-top-wrap .VirtualizedSelectFocusedOption,
      #ui-top-wrap .VirtualizedSelectSelectedOption {{
        height: 44px !important;
        min-height: 44px !important;
        font-size: 40px !important;
        line-height: 44px !important;
        padding: 0 10px !important;
        display: flex !important;
        align-items: center !important;
      }}
      /* Dropdown arrow for top bar */
      #ui-top-wrap .Select-arrow-zone {{
        position: absolute !important;
        right: 0 !important;
        top: 0 !important;
        bottom: 0 !important;
        width: 40px !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
      }}
      #ui-top-wrap .Select-arrow {{
        border-color: #555 transparent transparent !important;
        border-style: solid !important;
        border-width: 10px 8px 0 !important;
      }}

      /* Top bar popover text/number inputs - same as right panel */
      #ui-top-wrap input[type="number"] {{
        height: 64px !important;
        padding: 6px 10px !important;
        border: 2px solid #A0A0A0 !important;
        border-radius: 6px !important;
        background: #FFFFFF !important;
        box-shadow: none !important;
        font-size: 40px !important;
        -webkit-appearance: textfield !important;
        -moz-appearance: textfield !important;
        appearance: textfield !important;
      }}
      #ui-top-wrap input[type="number"]:focus {{
        outline: none !important;
        border-color: #707070 !important;
      }}
      #ui-top-wrap input[type="number"]::-webkit-inner-spin-button,
      #ui-top-wrap input[type="number"]::-webkit-outer-spin-button {{
        -webkit-appearance: none !important;
        margin: 0 !important;
        display: none !important;
      }}

      /* Download popover: fixed control widths (reuse right-panel custom number spinner) */
      #download-panel button.spinner-btn {{
        height: auto !important;
        min-height: 0 !important;
      }}
      #download-panel input[type="number"] {{
        box-sizing: border-box !important;
        width: 300px !important;
        min-width: 300px !important;
        max-width: 300px !important;
      }}
      #download-panel button.spinner-btn {{
        height: auto !important;
        min-height: 0 !important;
      }}
      #download-panel input[type="text"] {{
        box-sizing: border-box !important;
        width: 300px !important;
        min-width: 300px !important;
        max-width: 300px !important;
      }}
      #download-panel .Select,
      #download-panel .Select-control {{
        box-sizing: border-box !important;
        width: 300px !important;
        min-width: 300px !important;
        max-width: 300px !important;
      }}

      /* Ordering popover: fixed control widths */
      #ordering-panel input[type="number"] {{
        box-sizing: border-box !important;
        width: 300px !important;
        min-width: 300px !important;
        max-width: 300px !important;
      }}


      #ui-top-wrap input[type="text"] {{
        height: 64px !important;
        padding: 6px 10px !important;
        border: 2px solid #A0A0A0 !important;
        border-radius: 6px !important;
        background: #FFFFFF !important;
        box-shadow: none !important;
        font-size: 40px !important;
      }}
      #ui-top-wrap input[type="text"]:focus {{
        outline: none !important;
        border-color: #707070 !important;
      }}

      /* Top bar button styling */
      #ui-top-wrap button {{
        height: 64px !important;
        width: 240px !important;
        font-size: 40px !important;
        border: 2px solid #A0A0A0 !important;
        border-radius: 6px !important;
        background: #F2F2F2 !important;
        box-sizing: border-box !important;
        cursor: pointer !important;
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
      }}
      #ui-top-wrap button:hover {{
        background: #E8E8E8 !important;
      }}

      /* Spinner buttons inside number inputs should not inherit top-bar button styling */
      #ui-top-wrap button.spinner-btn {{
        height: auto !important;
        width: auto !important;
        font-size: 10px !important;
        border: none !important;
        border-radius: 0px !important;
        background: transparent !important;
        padding: 0px !important;
        margin: 0px !important;
        box-sizing: border-box !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        line-height: 1 !important;
        cursor: pointer !important;
      }}
      #ui-top-wrap button.spinner-btn:hover {{
        background: transparent !important;
      }}

      /* Download popover confirm button should span label-to-control width */
      #ui-top-wrap #btn-download-confirm {{
        width: 640px !important;
        white-space: nowrap !important;
      }}

      /* Ordering popover confirm button should span label-to-control width */
      #ui-top-wrap #btn-ordering-confirm {{
        width: 840px !important;
        white-space: nowrap !important;
      }}

      /* Top-bar: all direct children stretch to 64px container height */
      #top-bar button {{
        min-width: 120px !important;
        height: 64px !important;
        white-space: nowrap !important;
        box-sizing: border-box !important;
        font-size: 40px !important;
        vertical-align: top !important;
      }}
      /* Upload wrapper: explicit 64px */
      #upload-slices {{
        display: inline-flex !important;
        align-items: stretch !important;
        height: 64px !important;
        vertical-align: top !important;
      }}
      #upload-slices button {{
        min-width: 120px !important;
        height: 64px !important;
        width: 100% !important;
        white-space: nowrap !important;
        box-sizing: border-box !important;
        font-size: 40px !important;
      }}
      /* ── Custom loading overlay (replaces dcc.Loading) ──────────────
         Only heavy callbacks activate it via running=.
         animation-delay = threshold: spinner is invisible for fast ops. */
      .loading-overlay {{
        position: absolute;
        inset: 0;
        display: flex;
        align-items: center;
        justify-content: center;
        background: rgba(255,255,255,0.55);
        z-index: 50;
        pointer-events: none;
        opacity: 0;
        visibility: hidden;
      }}
      .loading-overlay.active {{
        pointer-events: auto;
        animation: overlay-appear 0.18s ease-out 0.55s both;
      }}
      @keyframes overlay-appear {{
        from {{ opacity: 0; visibility: visible; }}
        to   {{ opacity: 1; visibility: visible; }}
      }}
      .loading-spinner {{
        /* vw is DPR-independent: same physical size on all browsers.
           2.5vw of canvas area ≈ 40px on a typical 1600px canvas.
           Using separate border properties avoids calc()-in-shorthand issues. */
        width: 2.5vw;
        height: 2.5vw;
        border-width: 0.24vw;
        border-style: solid;
        border-color: #DDD;
        border-top-color: #555;
        border-radius: 50%;
        animation: spin 0.75s linear infinite;
      }}
      @keyframes spin {{
        to {{ transform: rotate(360deg); }}
      }}

      /* ── Top-bar "Updating" status badge ──────────────────────────── */
      /* Hidden by default; clientside callback flips display:flex when any
         of the loading flags is true, or while we're waiting for Plotly
         to finish a render. Sized to match toolbar buttons (40px font,
         64px height) so it visually integrates with the row. */
      #updating-status {{
        display: none;                 /* JS toggles to 'flex' */
        align-items: center;
        margin-left: 10px;             /* same gap as other toolbar buttons */
        height: 64px;
        padding: 0 24px;
        border-radius: 32px;           /* pill shape (= height / 2) */
        background-color: #FFF4D6;
        color: #7A5A00;
        font-size: 40px;
        font-weight: 500;
        white-space: nowrap;
        user-select: none;
        box-sizing: border-box;
      }}
      #updating-status .updating-dot {{
        display: inline-block;
        margin-right: 14px;
        color: #F08C00;
        animation: pulse-dot 1.2s ease-in-out infinite;
      }}
      @keyframes pulse-dot {{
        0%, 100% {{ opacity: 0.35; }}
        50%      {{ opacity: 1.0;  }}
      }}

      /* ── File order modal ──────────────────────────────────────── */
      #file-order-modal {{
        display: none;
        position: fixed;
        inset: 0;
        z-index: 99000;
        background: rgba(0,0,0,0.35);
        align-items: center;
        justify-content: center;
      }}
      #file-order-modal.visible {{
        display: flex !important;
      }}
      /* Every size uses calc(var(--dlg-base) * N) to avoid em compounding,
         which would make nested elements 2–3× larger than intended.
         dlg-base ≈ 0.37vw at 1920px ≈ 7.1px, scaled to match the 0.35
         density factor used elsewhere (was 1.05vw ≈ 20px, which made
         the dialog read as oversized after the rest of the UI was
         scaled down). vw is DPR-independent: same physical size on all
         browsers. */
      #file-order-dialog {{
        --dlg-base: clamp(4px, 0.37vw, 8.5px);
        background: #fff;
        border-radius: calc(var(--dlg-base) * 0.6);
        padding: calc(var(--dlg-base) * 1.8) calc(var(--dlg-base) * 2.0)
                 calc(var(--dlg-base) * 1.6) calc(var(--dlg-base) * 2.0);
        width: max-content;
        min-width: max-content;
        max-width: 90vw;
        box-shadow: 0 8px 32px rgba(0,0,0,0.22);
        font-family: inherit;
        font-size: var(--dlg-base);
        position: relative;
      }}
      #file-order-dialog h3 {{
        margin: 0 0 calc(var(--dlg-base) * 0.4) 0;
        font-size: calc(var(--dlg-base) * 2);
        font-weight: 600;
        color: #222;
      }}
      #file-order-dialog .subtitle {{
        font-size: calc(var(--dlg-base) * 1.6);
        color: #666;
        margin-bottom: calc(var(--dlg-base) * 1.0);
        white-space: nowrap;
      }}
      #file-order-dialog .order-row {{
        display: flex;
        align-items: center;
        padding: calc(var(--dlg-base) * 0.5) calc(var(--dlg-base) * 0.6);
        border-radius: calc(var(--dlg-base) * 0.3);
        font-size: calc(var(--dlg-base) * 1.7);
        background: #FAFAFA;
        border: 2px solid #E0E0E0;
        margin-bottom: calc(var(--dlg-base) * 0.4);
      }}
      #file-order-dialog .order-row:hover {{
        background: #F0F4FF;
      }}
      #file-order-dialog .order-seq {{
        width: calc(var(--dlg-base) * 1.9);
        text-align: right;
        color: #888;
        font-size: calc(var(--dlg-base) * 1.4);
        flex-shrink: 0;
        margin-right: calc(var(--dlg-base) * 0.5);
      }}
      #file-order-dialog .order-name {{
        flex: 1 1 auto;
        white-space: nowrap;
        color: #222;
        padding-right: calc(var(--dlg-base) * 0.4);
      }}
      #file-order-dialog .order-arrow-btn {{
        width: calc(var(--dlg-base) * 2.0) !important;
        height: calc(var(--dlg-base) * 2.0) !important;
        min-width: calc(var(--dlg-base) * 2.0) !important;
        padding: 0 !important;
        font-size: calc(var(--dlg-base) * 1.3) !important;
        line-height: calc(var(--dlg-base) * 2.0) !important;
        text-align: center !important;
        border-radius: calc(var(--dlg-base) * 0.25) !important;
        border: 2px solid #CCC !important;
        background: #FFF !important;
        cursor: pointer !important;
        flex-shrink: 0 !important;
        margin-left: calc(var(--dlg-base) * 0.2) !important;
        color: #444 !important;
        font-family: inherit !important;
      }}
      #file-order-dialog .order-arrow-btn:hover {{
        background: #E8EAF6 !important;
        border-color: #9FA8DA !important;
      }}
      #file-order-dialog .order-arrow-btn:disabled {{
        opacity: 0.3 !important;
        cursor: default !important;
      }}
      #file-order-actions {{
        display: flex;
        justify-content: flex-end;
        gap: calc(var(--dlg-base) * 0.7);
        margin-top: calc(var(--dlg-base) * 1.0);
      }}
      /* Both buttons share an explicit height + box-sizing so the
         1.5px border on Cancel vs the borderless Generate cannot push
         them to different visual heights. Without this, browsers'
         built-in <button> padding compounds asymmetrically with the
         border-vs-no-border styles, producing the "tall Cancel /
         flat Generate" mismatch. */
      #btn-order-cancel,
      #btn-order-generate {{
        box-sizing: border-box !important;
        height: calc(var(--dlg-base) * 3.2) !important;
        line-height: 1 !important;
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
      }}
      #btn-order-cancel {{
        font-size: calc(var(--dlg-base) * 1.8) !important;
        padding: 0 calc(var(--dlg-base) * 1.4) !important;
        border-radius: calc(var(--dlg-base) * 0.3) !important;
        border: 1.5px solid #BBBBBB !important;
        background: #FFF !important;
        color: #444 !important;
        cursor: pointer !important;
        font-family: inherit !important;
      }}
      #btn-order-generate {{
        font-size: calc(var(--dlg-base) * 1.8) !important;
        padding: 0 calc(var(--dlg-base) * 1.4) !important;
        border-radius: calc(var(--dlg-base) * 0.3) !important;
        border: none !important;
        background: #1a73e8 !important;
        color: #FFF !important;
        cursor: pointer !important;
        font-weight: 600 !important;
        font-family: inherit !important;
      }}
      #btn-order-generate:hover {{
        background: #1558c0 !important;
      }}
      #btn-order-cancel:hover {{
        background: #F0F0F0 !important;
      }}
      #file-order-list {{
        max-height: calc(var(--dlg-base) * 15 * 3.1);
        overflow-y: auto;
        overflow-x: hidden;
        margin-bottom: calc(var(--dlg-base) * 0.2);
      }}
      #file-order-upload-hint {{
        font-size: calc(var(--dlg-base) * 1.5);
        color: #888;
        margin-bottom: calc(var(--dlg-base) * 0.5);
        min-height: calc(var(--dlg-base) * 1.5);
      }}

      /* category-combo: explicit 64px */
      .category-combo-btn {{
        height: 64px !important;
        box-sizing: border-box !important;
      }}
      .category-combo-label {{
        height: 64px !important;
        box-sizing: border-box !important;
      }}

      /* Category combo dropdown overrides - use ID selector for higher specificity */
      #ui-top-wrap #category-col .Select-control {{
        border: 0px !important;
        border-radius: 0px !important;
        background: #FFFFFF !important;
        box-shadow: none !important;
        display: flex !important;
        align-items: center !important;
        height: 64px !important;
      }}
      #ui-top-wrap #category-col .Select-multi-value-wrapper {{
        display: flex !important;
        align-items: center !important;
        flex: 1 !important;
        height: 64px !important;
      }}
      #ui-top-wrap #category-col .Select-value {{
        position: relative !important;
        top: auto !important;
        transform: none !important;
        display: flex !important;
        align-items: center !important;
        height: 64px !important;
        line-height: 64px !important;
        padding-left: 14px !important;
        /* padding-right bumped 10→24px so the selected value text
           ("NAICS Industry" etc.) doesn't sit right up against the
           dropdown arrow — previous 10px looked visually cramped. */
        padding-right: 24px !important;
      }}
      #ui-top-wrap #category-col .Select-value-label {{
        font-size: 40px !important;
        line-height: 64px !important;
      }}
      #ui-top-wrap #category-col .Select-placeholder {{
        position: relative !important;
        top: auto !important;
        transform: none !important;
        display: flex !important;
        align-items: center !important;
        height: 64px !important;
        line-height: 64px !important;
        padding-left: 14px !important;
        padding-right: 24px !important;
      }}
      #ui-top-wrap #category-col .Select-input {{
        height: 64px !important;
        display: flex !important;
        align-items: center !important;
      }}
      #ui-top-wrap #category-col .Select-arrow-zone {{
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        /* 42→54px: gives the arrow more breathing room on both sides
           without nudging the label text any further left. */
        width: 54px !important;
        height: 64px !important;
        flex-shrink: 0 !important;
        padding-right: 8px !important;
      }}
      #ui-top-wrap #category-col .Select-arrow {{
        border-color: #555 transparent transparent !important;
        border-style: solid !important;
        border-width: 10px 8px 0 !important;
      }}


      /* Tooltip (hoverlayer) counter-scaling to keep fixed size */
      .js-plotly-plot .hoverlayer {{
        transform: none;
        transform-origin: 0 0;
      }}

      /* Prevent any transitions on hover elements to avoid size flash */
      .js-plotly-plot .hoverlayer,
      .js-plotly-plot .hoverlayer *,
      .js-plotly-plot .hovertext,
      .js-plotly-plot .hovertext * {{
        transition: none !important;
      }}

      /* Force hover-label font-size at the SVG level, expressed in vw.
         WHY (correctness): plotly's annotation hovertext renders through
         the same .hoverlayer/.hovertext SVG group as trace hover, but in
         some plotly.js versions the per-annotation `hoverlabel.font.size`
         setter is dropped between the layout dict and the SVG attrs
         applied to <text>, leaving annotation tooltips at the plotly
         default (~13px) while trace tooltips render at the layout-level
         hoverlabel size. Setting the size in CSS bypasses the JS path
         entirely and applies uniformly to both — which is also what we
         want product-wise: every hover tooltip should look identical.

         WHY (vw, not px): the rest of the app sizes UI relative to the
         viewport (right-panel uses `uiScale = innerWidth / (6*920)`,
         canvas hint text uses `0.75vw`) so visual proportions stay
         constant across browsers / display sizes. Using a fixed px
         value here would make tooltips disproportionately large on
         small viewports and disproportionately small on 4K displays,
         breaking the visual consistency the rest of the app maintains.

         CONVERSION: cfg.tooltip_font_size is treated as the *reference*
         px size at a 1920 px viewport (typical 1080p design width).
         At any viewport `w`, the displayed size becomes:
              cfg.tooltip_font_size × (w / 1920)
         Implemented as `vw` so the browser computes this automatically:
              vw_value = cfg.tooltip_font_size / 1920 × 100
         Example: cfg=40 → 2.083vw → 40px@1920, 53px@2560, 28px@1366. */
      .js-plotly-plot .hoverlayer .hovertext text,
      .js-plotly-plot .hoverlayer .hovertext text *,
      .js-plotly-plot .hoverlayer text {{
        /* Counter-scale via CSS variable.
           --canvas-zoom is set on :root from JS (applyCanvasTransform).
           Why this rather than a plain `... vw`: the .hoverlayer SVG sits
           INSIDE #canvas-content, which has `transform: scale(canvasZoom)`.
           A static vw value would be visually multiplied by canvasZoom, so
           tooltip text would grow/shrink with chart zoom — the bug we just
           hit. Dividing by --canvas-zoom here makes the rendered text
           visually constant: the CSS shrinks the font when canvas zooms in,
           the canvas transform scales it back up, net effect is a stable
           visual size. The `!important` is required because plotly sets an
           SVG `font-size` attribute on each <text> from
           _fullLayout.hoverlabel.font.size (which we don't touch here). */
        font-size: calc({(cfg.tooltip_font_size / 1920) * 100:.4f}vw / var(--canvas-zoom, 1)) !important;
      }}

      /* Hide plotly's NATIVE annotation hover tooltip.
         Setting `hovertext` on an annotation (which we do for truncated
         enclosure labels) makes plotly render its own SVG-based tooltip
         anchored to the annotation's (x, y). On rotated 90° group labels
         that anchor lands far from the visible text — top-left of the
         figure, gaps between slices, etc — looking broken. We instead
         render our own HTML overlay that follows the cursor (see the
         mousemove handler in app.index_string), and use this CSS to
         physically suppress plotly's broken native rendering. The
         underlying `hovertext` payload on the annotation object is
         untouched; only the SVG rendering is hidden. Trace tooltips
         live in a different SVG group and are unaffected. */
      .js-plotly-plot .annotation-hovertext,
      .js-plotly-plot g.hovertext.annotation-hovertext-g,
      .js-plotly-plot g[class*="annotation-hover"] {{
        display: none !important;
        visibility: hidden !important;
      }}

      /* Plotly inner SVG elements must overflow to avoid clipping tooltips/annotations.
         DESIGN NOTE: This deliberately relaxes clipping at every inner layer.
         Annotations (e.g. left-side slice labels) extend beyond main-svg's drawing
         rect and would be clipped otherwise. Final clipping is enforced by
         #graph-wrap {{ overflow: hidden }}. Do NOT remove that outer clip, or
         zoomed-in content will spill over the top-bar/right-panel. */
      .js-plotly-plot, .plot-container, .svg-container, .main-svg {{
        overflow: visible !important;
      }}
      .js-plotly-plot .hoverlayer {{
        overflow: visible !important;
      }}
      .js-plotly-plot .main-svg {{
        overflow: visible !important;
      }}
      .js-plotly-plot svg {{
        overflow: visible !important;
      }}
      #_dash-app-content, ._dash-loading, .dash-graph {{
        overflow: visible !important;
      }}

    </style>
</head>
<body>
    {{%app_entry%}}
    <footer>
        {{%config%}}
        {{%scripts%}}
        {{%renderer%}}
    </footer>
    <script>
    (function() {{
        // Wait for DOM to be ready
        function initCanvasControls() {{
            var graphWrap = document.getElementById('graph-wrap');
            var uiTopWrap = document.getElementById('ui-top-wrap');
            var rightPanel = document.getElementById('right-panel');
            var rightPanelInner = document.getElementById('right-panel-inner');
            var rightPanelSpacer = document.getElementById('right-panel-spacer');

            if (!graphWrap || !rightPanel || !uiTopWrap || !rightPanelInner || !rightPanelSpacer) {{
                setTimeout(initCanvasControls, 100);
                return;
            }}


            // ---- State variables ----
            // UI zoom state — set automatically on init so right panel = 1/6 viewport
            var _NATURAL_PANEL_W = 920;   // panel layout width the controls were designed for
            var uiMinScale = 0.15;         // lowered to support small viewports
            var uiMaxScale = 2.0;
            // Initial scale: visual panel = 1/6 of viewport width
            var uiScale = Math.min(uiMaxScale,
                           Math.max(uiMinScale,
                           window.innerWidth / (6 * _NATURAL_PANEL_W)));

            function getGraphElement() {{
                return document.getElementById('graph');
            }}

            function isInsideGraphWrap(element) {{
                while (element) {{
                    if (element === graphWrap) return true;
                    element = element.parentElement;
                }}
                return false;
            }}

            function isInsideUI(element) {{
                while (element) {{
                    if (element === uiTopWrap || element === rightPanel) return true;
                    element = element.parentElement;
                }}
                return false;
            }}

            function getPlotlyGd() {{
                var graph = getGraphElement();
                if (!graph) return null;
                var gd = graph.querySelector('.js-plotly-plot');
                return (gd && gd._fullLayout) ? gd : null;
            }}


            // ---- UI transform ----
            function applyUITransform() {{
                var PANEL_W = _NATURAL_PANEL_W;
                if (uiTopWrap) {{
                    uiTopWrap.style.transform = 'scale(' + uiScale + ')';
                    uiTopWrap.style.transformOrigin = 'top left';
                    uiTopWrap.style.width = (100 / uiScale) + '%';
                }}
                if (rightPanel) {{
                    rightPanel.style.transformOrigin = 'top right';
                    rightPanel.style.transform = 'scale(' + uiScale + ')';
                    rightPanel.style.width = PANEL_W + 'px';
                    // height: layout = 100/uiScale% so visual height = 100% of viewport
                    rightPanel.style.height = (100 / uiScale) + '%';
                    // No marginLeft needed — transformOrigin:top-right keeps the right
                    // edge fixed; the visual left edge moves to viewport_w - PANEL_W*uiScale.
                    rightPanel.style.marginLeft = '';
                }}
                if (rightPanelInner) {{
                    rightPanelInner.style.transform = '';
                    rightPanelInner.style.transformOrigin = '';
                    if (rightPanelSpacer) {{
                        rightPanelSpacer.style.height = '0px';
                    }}
                }}
                // ── Chart stays fixed when panel scales ──────────────────────
                // Earlier this block updated graphWrap.style.width and called
                // Plotly.Plots.resize() on every uiScale change, so the chart
                // would expand/shrink in lockstep with the panel. That caused
                // a visual mismatch: chart container resized (data-unit blocks
                // got larger or smaller) but absolute-px values like block
                // text font size, border widths, line widths in
                // base_sizes.traces did NOT change — leaving text the same
                // size in a now-different-sized block. With small uiScale
                // (panels shrunk), blocks expanded but text stayed at cfg
                // size and overflowed; with large uiScale, blocks shrunk
                // but text stayed full size and looked oversized.
                //
                // User preference: chart should stay visually CONSTANT when
                // panels scale. Decoupling the two means the chart's
                // physical width is whatever it was on initial layout, and
                // panel scaling only affects the panel itself. The panel's
                // CSS transform: scale(uiScale) is applied at the panel
                // root with transform-origin:top-right, so the visible
                // panel area shrinks toward / grows from the right edge
                // — leaving the empty space to its left, which the chart
                // does NOT auto-claim. This keeps text-vs-block ratios
                // stable across all panel sizes.
                //
                // graphWrap.style.width is left alone (set by the initial
                // CSS rule, not overwritten per uiScale change), and
                // Plotly.Plots.resize() is not invoked on UI transforms.
                // Keep canvas hints readable on any viewport — use 0.75vw as base size.
                // This is independent of uiScale so they don't become tiny on small screens.
                var _hintPx = (window.innerWidth * 0.0075).toFixed(2) + 'px';
                var zoomHint = document.getElementById('canvas-zoom-hint');
                if (zoomHint) zoomHint.style.fontSize = _hintPx;
                // Re-sync top/height for graph-wrap (uses offsetHeight × uiScale).
                _lastSyncH = -1;
                syncRightPanelTop();
                // Do NOT call Plotly.Plots.resize() here. UI scaling is
                // independent of chart sizing — see comment above about
                // user preference for chart staying visually constant
                // when panels are scaled.

                // Modal dialog & spinner now use pure CSS vw sizing — no JS scaling needed.
            }}

            // =====================================================
            // Canvas zoom + pan via CSS transform (image-style, zero lag)
            // No Plotly axis recalculation — everything scales as one unit.
            // =====================================================
            window._applyUITransform = applyUITransform;   // expose for modal-open callback
            var canvasZoom = 1.0;
            var canvasTx   = 0.0;
            var canvasTy    = 0.0;
            var canvasMinZoom = 0.1;
            var canvasMaxZoom = 20.0;

            function getCanvasContent() {{
                return document.getElementById('canvas-content');
            }}

            function applyCanvasTransform() {{
                var cc = getCanvasContent();
                if (!cc) return;
                cc.style.transform =
                    'translate(' + canvasTx.toFixed(2) + 'px,' + canvasTy.toFixed(2) + 'px)' +
                    ' scale(' + canvasZoom + ')';

                // Publish canvasZoom as a CSS custom property on :root so
                // CSS rules can counter-scale things that live inside
                // #canvas-content. Used by .hoverlayer text font-size to
                // keep block/band tooltips visually constant across zoom
                // (calc(... / var(--canvas-zoom)) cancels the scale()).
                // Must be a plain number string (NOT including 'px'/'vw')
                // so it can be used as a divisor in calc().
                document.documentElement.style.setProperty(
                    '--canvas-zoom', String(canvasZoom)
                );

                // ── Hover tooltip font: dual-track sync ───────────────────────
                // Two parallel tracks must agree on the same px value, or
                // the user sees flicker during ctrl-wheel zoom:
                //
                //   Track 1 (CSS):  .hoverlayer text font-size is computed
                //   from `calc((cfg.tooltip_font_size/1920)*100vw / var(--canvas-zoom))`.
                //   Updates immediately when --canvas-zoom changes (set just
                //   above).
                //
                //   Track 2 (SVG attribute):  plotly sets `font-size` as an
                //   SVG attribute on each <text> in the hoverlayer, sourced
                //   from `_fullLayout.hoverlabel.font.size`. CSS `!important`
                //   wins for rendering, but during plotly's hoverlabel
                //   rebuild (which happens on every hover state change —
                //   including the ones triggered indirectly by ctrl-wheel
                //   canvas zoom) plotly uses the SVG attribute to size the
                //   background <rect> via getBBox(). If the attribute is
                //   stale (e.g., cfg.tooltip_font_size = 18 when CSS calc
                //   says 12 at canvasZoom=1.5), the rect snaps to 18-sized
                //   bbox for one frame, then CSS recomputes to 12 →
                //   "one shrink per wheel tick" — exactly the symptom.
                //
                // Keeping _fullLayout.hoverlabel.font.size in sync with the
                // CSS calc value makes the rect computation match what CSS
                // ultimately renders. No frame-disagreement → no flicker.
                // Direct mutation does NOT trigger a plotly re-render
                // (plotly has no reactive watcher on _fullLayout fields);
                // it only takes effect on the next natural hover event.
                var TT_VW_FRAC = {(cfg.tooltip_font_size / 1920):.6f};
                var hoverFontPx = Math.max(6, window.innerWidth * TT_VW_FRAC / canvasZoom);
                var gdh = getPlotlyGd();
                if (gdh) {{
                    try {{ gdh._fullLayout.hoverlabel.font.size = hoverFontPx; }} catch(e) {{}}
                    try {{
                        if (!gdh.layout.hoverlabel) gdh.layout.hoverlabel = {{}};
                        if (!gdh.layout.hoverlabel.font) gdh.layout.hoverlabel.font = {{}};
                        gdh.layout.hoverlabel.font.size = hoverFontPx;
                    }} catch(e) {{}}
                }}

                // ── No-data HTML hint: size + counter-scale only ─────────────
                // Visibility is controlled exclusively by the Python toggle_no_data_hint
                // callback.  JS only adjusts font-size and transform so it stays
                // centred and readable regardless of canvasZoom.
                var hint = document.getElementById('no-data-hint-html');
                if (hint && hint.style.display !== 'none') {{
                    var hintPx = (window.innerWidth * 0.0075 / canvasZoom).toFixed(2);
                    hint.style.fontSize = hintPx + 'px';
                    hint.style.transform =
                        'translate(-50%,-50%) scale(' + (1 / canvasZoom).toFixed(6) + ')';
                }}
            }}

            function resetCanvasTransform() {{
                canvasZoom = 1.0; canvasTx = 0.0; canvasTy = 0.0;
                _containFitKey = null;   // allow next figure load to re-fit
                applyCanvasTransform();
            }}

            // ── Wheel handler ─────────────────────────────────────────
            document.addEventListener('wheel', function(e) {{
                if (e.ctrlKey) {{
                    e.preventDefault();
                    if (isInsideGraphWrap(e.target)) {{
                        // Canvas zoom: scale about the cursor position
                        var factor = e.deltaY > 0 ? (1.0 / 1.1) : 1.1;
                        var newZoom = Math.min(canvasMaxZoom, Math.max(canvasMinZoom, canvasZoom * factor));
                        factor = newZoom / canvasZoom;   // actual factor after clamping
                        canvasZoom = newZoom;
                        var rect = graphWrap.getBoundingClientRect();
                        var mx = e.clientX - rect.left;
                        var my = e.clientY - rect.top;
                        // Keep the point under the cursor stationary
                        canvasTx = mx - (mx - canvasTx) * factor;
                        canvasTy  = my - (my - canvasTy)  * factor;
                        applyCanvasTransform();
                        e.stopPropagation();
                    }} else if (isInsideUI(e.target)) {{
                        e.stopPropagation();
                        var delta = e.deltaY > 0 ? -0.05 : 0.05;
                        var newScale = Math.min(uiMaxScale, Math.max(uiMinScale, uiScale + delta));
                        if (newScale !== uiScale) {{
                            uiScale = newScale;
                            applyUITransform();
                        }}
                    }} else {{
                        e.stopPropagation();
                    }}
                }} else {{
                    if (isInsideGraphWrap(e.target)) {{
                        e.stopPropagation();
                        e.preventDefault();
                    }}
                }}
            }}, {{ passive: false, capture: true }});

            // ── Middle-drag pan ───────────────────────────────────────
            var _isPanning = false, _panLastX = 0, _panLastY = 0;
            document.addEventListener('mousedown', function(e) {{
                if (e.button === 1 && isInsideGraphWrap(e.target)) {{
                    _isPanning = true;
                    _panLastX = e.clientX;
                    _panLastY = e.clientY;
                    e.preventDefault();
                }}
            }});
            document.addEventListener('mousemove', function(e) {{
                if (!_isPanning) return;
                canvasTx += e.clientX - _panLastX;
                canvasTy  += e.clientY - _panLastY;
                _panLastX = e.clientX;
                _panLastY = e.clientY;
                applyCanvasTransform();
            }});
            document.addEventListener('mouseup', function(e) {{
                if (e.button === 1) _isPanning = false;
            }});

            // Double-click on canvas: reset zoom/pan to default
            graphWrap.addEventListener('dblclick', function(e) {{
                if (!e.ctrlKey) resetCanvasTransform();
            }});
            // Left-click on graph wrap: signal Dash for highlight-clear logic.
            // We need to fire btn-wrap-leftclick in TWO cases:
            //   (a) click was outside #graph entirely (graph-wrap padding,
            //       no-data hint, etc.)
            //   (b) click was inside #graph but did NOT hit any trace (i.e.,
            //       it landed on the canvas background between blocks/bands).
            //       In this case Plotly does NOT fire clickData, so without
            //       wrap-click the user's "click-to-deselect" intent is lost.
            //
            // We must NOT fire wrap-click when the click HIT a trace, because
            // that creates a race with Plotly's clickData → trig_clickdata
            // path: a stray trig_wrap-only invocation arriving second would
            // clear the just-set highlight (the nonce defense in
            // render_with_highlight cannot reliably distinguish "just set"
            // from prior state because nonce_now and nonce_prev both read
            // the same store-click-nonce State).
            //
            // Detection: walk up from e.target. If we find an ancestor with
            // class `trace` (Plotly tags every trace's <g> container with
            // this class), the click hit a trace and we must NOT fire.
            // Otherwise (background click), we fire.
            graphWrap.addEventListener('click', function(e) {{
                if (e.button !== 0) return;
                var hitTrace = e.target.closest && e.target.closest('.trace');
                if (hitTrace) {{
                    // clickData will handle it.
                    return;
                }}
                var btn = document.getElementById('btn-wrap-leftclick');
                if (btn) btn.click();
            }});

            // Prevent context menu on middle click
            graphWrap.addEventListener('auxclick', function(e) {{
                if (e.button === 1) e.preventDefault();
            }});

            // =====================================================
            // File-order modal: click backdrop to close (event delegation)
            // =====================================================
            document.addEventListener('click', function(e) {{
                if (e.target && e.target.id === 'file-order-modal') {{
                    var cancelBtn = document.getElementById('btn-order-cancel');
                    if (cancelBtn) cancelBtn.click();
                }}
            }});

            // Ctrl+Double-click on UI to reset UI scale
            function handleUIDoubleClick(e) {{
                if (e.ctrlKey) {{
                    e.preventDefault();
                    uiScale = 1.0;
                    applyUITransform();
                }}
            }}
            if (uiTopWrap) uiTopWrap.addEventListener('dblclick', handleUIDoubleClick);
            if (rightPanel) rightPanel.addEventListener('dblclick', handleUIDoubleClick);

            // =====================================================
            // Initial "contain fit": on first load of a new figure,
            // expand the looser axis so both axes share the same
            // px-per-data-unit — diagram fits in both dimensions
            // with a small visual margin.  Flag prevents re-entry.
            // =====================================================
            var _containFitKey  = null;   // meta.init_x_span of last fitted figure
            var _containFitDone = false;

            function applyContainFit(gd) {{
                if (_containFitDone) return;
                var meta = gd._fullLayout && gd._fullLayout.meta;
                if (!meta || !meta.init_x_span) return;

                var key = meta.init_x_span;
                if (key === _containFitKey) return;   // already fitted for this figure
                _containFitKey  = key;
                _containFitDone = true;
                setTimeout(function() {{ _containFitDone = false; }}, 400);

                var cw = graphWrap.offsetWidth;
                var ch = graphWrap.offsetHeight;
                if (cw <= 0 || ch <= 0) return;

                var xax = gd._fullLayout.xaxis;
                var yax = gd._fullLayout.yaxis;
                if (!xax || !xax.range || !yax || !yax.range) return;

                var xs = Math.abs(xax.range[1] - xax.range[0]);
                var ys = Math.abs(yax.range[1] - yax.range[0]);
                if (xs <= 0 || ys <= 0) return;

                // px-per-data-unit in each direction at canvasZoom=1
                var ppu_x = cw / xs;
                var ppu_y = ch / ys;

                // "Fill which fills first": the tighter direction (smaller ppu)
                // limits the zoom. Scale canvasZoom so the tighter direction
                // occupies (1 - 2*PAD) of the canvas; the looser direction
                // then occupies < (1 - 2*PAD) — it has more breathing room.
                var PAD = 0.06;   // 6% margin each side
                // At canvasZoom=1 each direction already fills 100%.
                // We want the tighter direction to fill (1-2*PAD).
                // Since CSS canvasZoom scales both uniformly:
                //   zoom × ppu_tight × data_tight / canvas_tight = (1 - 2*PAD)
                //   but ppu_tight × data_tight = canvas_tight → zoom = (1 - 2*PAD)
                // …which is just a uniform shrink.  Apply it:
                var ppu_tight = Math.min(ppu_x, ppu_y);
                // zoom = fraction of canvas the tighter direction should occupy
                var zoom = 1.0 - 2.0 * PAD;   // = 0.88

                // Extra safety: if looser direction is much larger (wide/narrow diagram),
                // scale down further so the actual DATA extent fits within the canvas.
                // Looser ppu fills: data_loose * ppu_loose / canvas_loose at zoom=1 = 1
                // — already fills. zoom just applies uniform shrink. ✓

                // Clamp 50%–100%
                zoom = Math.min(1.0, Math.max(0.5, zoom));

                // Center: with transformOrigin 0 0, after scale(zoom) the canvas-content
                // occupies zoom×cw × zoom×ch anchored at (0,0). Shift to centre it.
                var tx = (cw * (1.0 - zoom)) * 0.5;
                var ty = (ch * (1.0 - zoom)) * 0.5;

                canvasZoom = zoom;
                canvasTx   = tx;
                canvasTy    = ty;
                applyCanvasTransform();
            }}

            // ── Truncated label tooltips (custom HTML div with explicit size) ──
            // Strategy: inject data-hovertext attribute on each annotation text node
            // via textContent matching (reliable, index-free), then use a floating
            // HTML div for the tooltip so we can control font size precisely.
            (function() {{
                var _ttDiv = null;
                function _getTtDiv() {{
                    if (!_ttDiv) {{
                        _ttDiv = document.createElement('div');
                        _ttDiv.id = '_lbl_tt';
                        // Font size sourced from cfg.tooltip_font_size — the
                        // SAME value the layout's plotly `hoverlabel` uses
                        // for block tooltips (see fig.update_layout above).
                        // Both stay in lockstep when cfg.tooltip_font_size
                        // is changed; both are fixed pixel sizes that do
                        // not scale with figure zoom.
                        //
                        // We set most properties via cssText (cheap, single
                        // assignment) but font-size goes through
                        // setProperty(..., 'important') as a separate step.
                        // Older browsers and certain Plotly stylesheets have
                        // historically dropped !important inside cssText
                        // batches, leaving font-size at the user-agent
                        // default — which is exactly the symptom we are
                        // fixing here.
                        // Font size: same vw expression we use for plotly's
                        // hover-text SVG (see CSS rule .js-plotly-plot
                        // .hoverlayer text). cfg.tooltip_font_size is
                        // interpreted as the px size at a 1920-px reference
                        // viewport; rendering as vw lets the tooltip scale
                        // with the browser width to match the rest of the
                        // app's viewport-relative UI sizing strategy
                        // (uiScale, canvas-hint 0.75vw, etc.).
                        _ttDiv.style.cssText = [
                            'position:fixed',
                            'background:#2b2b2b',
                            'color:#ffffff',
                            'padding:8px 12px',
                            'border-radius:4px',
                            'font-family:Arial,sans-serif',
                            'font-weight:400',
                            'line-height:1.35',
                            'pointer-events:none',
                            'display:none',
                            'z-index:999999',
                            'white-space:pre-wrap',
                            'max-width:480px',
                            'box-shadow:2px 3px 10px rgba(0,0,0,0.5)'
                        ].join(';');
                        _ttDiv.style.setProperty('font-size', '{(cfg.tooltip_font_size / 1920) * 100:.4f}vw', 'important');
                        document.body.appendChild(_ttDiv);
                    }}
                    return _ttDiv;
                }}

                var _currentGd = null;

                function _plotlyTextMatches(domText, annoText) {{
                    var a = (domText || '').replace(/\\s+/g, '').trim();
                    var b = (annoText || '').replace(/<[^>]+>/g, '').replace(/\\s+/g, '').trim();
                    return a === b && a.length > 0;
                }}

                window._attachTruncatedLabelTooltips = function(gd) {{
                    _currentGd = gd;
                    try {{
                        var annos = (gd._fullLayout && gd._fullLayout.annotations) || [];
                        var textNodes = gd.querySelectorAll('.annotationlayer text');
                        for (var j = 0; j < textNodes.length; j++) {{
                            textNodes[j].removeAttribute('data-hovertext');
                        }}
                        var claimed = new Array(textNodes.length);
                        // Iterate annotations — those with `.hovertext` set
                        // are the truncated/wrapped ones (see Python helper
                        // _apply_truncation_names). Use textContent matching
                        // to claim the corresponding SVG <text> node.
                        for (var i = 0; i < annos.length; i++) {{
                            var a = annos[i];
                            if (!a) continue;
                            var fullText = a.hovertext;
                            if (!fullText) continue;
                            for (var j2 = 0; j2 < textNodes.length; j2++) {{
                                if (claimed[j2]) continue;
                                if (_plotlyTextMatches(textNodes[j2].textContent, a.text)) {{
                                    claimed[j2] = true;
                                    textNodes[j2].setAttribute('data-hovertext', fullText);
                                    break;
                                }}
                            }}
                        }}
                    }} catch (err) {{ console.warn('tt attach:', err); }}
                }};

                window._dumpTT = function() {{
                    if (!_currentGd) {{ console.log('no gd'); return; }}
                    var decorated = _currentGd.querySelectorAll(
                        '.annotationlayer text[data-hovertext]');
                    console.log('Decorated:', decorated.length);
                    decorated.forEach(function(el) {{
                        console.log('  visible="' + el.textContent.trim() +
                                    '", full="' + el.getAttribute('data-hovertext') + '"');
                    }});
                    var tt = document.getElementById('_lbl_tt');
                    if (tt) {{
                        var cs = window.getComputedStyle(tt);
                        console.log('tooltip computed font-size:', cs.fontSize);
                    }}
                    return decorated;
                }};

                // ── Cached hit-zone system (decouples hit detection
                //    from realtime SVG state) ─────────────────────────
                //
                // Plotly's hover handlers can mutate the annotation
                // <text> element's state (display / visibility /
                // transform) when the cursor enters a label. That
                // makes getBoundingClientRect on the same node return
                // 0×0 or wildly different coordinates the next frame —
                // tooltip flickers / disappears.
                //
                // Solution: snapshot every annotation label's bbox
                // ONCE per plotly redraw, when SVG state is fresh and
                // no hover is in progress. Cache that snapshot. On
                // mousemove, hit-test against the cache, NOT against
                // live DOM. Cache is invalidated on:
                //   • plotly_afterplot (hooked from tryAttachAfterPlot)
                //   • window scroll / resize (viewport coords change)
                //   • wheel events (canvas pan/zoom shifts labels)
                //
                // First mousemove after invalidation lazily rebuilds.

                var __ttHitZones = null;     // null = needs rebuild

                function _buildHitZones() {{
                    // Build hit zones from DATA COORDINATES, not SVG bboxes.
                    //
                    // This is the same approach plotly's own trace hover
                    // (block tooltip) uses internally: it knows each
                    // trace's data-space coordinates and converts to
                    // pixel coordinates via xaxis.l2p() / yaxis.l2p(),
                    // then hit-tests against those pixel rectangles.
                    // Result: rock-solid, completely independent of SVG
                    // DOM state. Plotly's hover-time mutations of <text>
                    // elements (display/visibility/transform changes)
                    // have ZERO effect on the hit zones we compute here.
                    //
                    // For each annotation that has a truncated_full_text
                    // entry we:
                    //   (1) read its data coordinates a.x, a.y
                    //   (2) convert to plot-area pixels via xaxis.l2p(a.x)
                    //   (3) offset by the plot area's screen position
                    //       (chartRect + size.l/.t margins)
                    //   (4) estimate text dimensions from font.size and
                    //       string length
                    //   (5) account for textangle rotation by swapping
                    //       width/height when |angle| > 45°
                    //   (6) add generous padding for hit tolerance
                    var zones = [];
                    var charts = document.querySelectorAll('.js-plotly-plot');

                    for (var ci = 0; ci < charts.length; ci++) {{
                        var chart = charts[ci];
                        var fl = chart._fullLayout;
                        if (!fl || !fl.annotations || !fl._size) continue;
                        var truncMap = (fl.meta && fl.meta.truncated_full_text) || {{}};
                        if (Object.keys(truncMap).length === 0) continue;

                        var xa = fl.xaxis, ya = fl.yaxis;
                        if (!xa || !ya) continue;
                        if (typeof xa.l2p !== 'function' || typeof ya.l2p !== 'function') continue;

                        var size = fl._size;
                        var chartRect = chart.getBoundingClientRect();
                        if (chartRect.width < 1 || chartRect.height < 1) continue;

                        // CSS transform scaling (canvas-content has
                        // translate+scale, so chartRect can differ from
                        // fl.width/.height by the canvasZoom factor).
                        var sx = (fl.width  > 0) ? (chartRect.width  / fl.width)  : 1;
                        var sy = (fl.height > 0) ? (chartRect.height / fl.height) : 1;

                        for (var ai = 0; ai < fl.annotations.length; ai++) {{
                            var a = fl.annotations[ai];
                            if (!a) continue;
                            var fullText = truncMap[String(ai)];
                            if (!fullText) continue;

                            // Only support data-space annotations (xref="x" / yref="y").
                            // Aggregation labels in this app all use data refs.
                            if (a.xref !== 'x' || a.yref !== 'y') continue;

                            // Data → chart-internal px (relative to plot area).
                            var px_in;
                            var py_in;
                            try {{
                                px_in = size.l + xa.l2p(a.x);
                                py_in = size.t + ya.l2p(a.y);
                            }} catch (err) {{ continue; }}
                            if (!isFinite(px_in) || !isFinite(py_in)) continue;

                            // Chart-internal → viewport (apply canvas-content scale).
                            var anchor_vx = chartRect.left + px_in * sx;
                            var anchor_vy = chartRect.top  + py_in * sy;

                            // Estimate text dimensions (in viewport px).
                            var fontSize = (a.font && a.font.size) || 14;
                            // Strip <br> and other tags before measuring.
                            // For multi-line text, use the longest line +
                            // count line breaks for height.
                            var textRaw = String(a.text || '');
                            var lines = textRaw.split(/<br\\s*\\/?>/i);
                            var maxLineLen = 0;
                            for (var li = 0; li < lines.length; li++) {{
                                var clean = lines[li].replace(/<[^>]+>/g, '');
                                if (clean.length > maxLineLen) maxLineLen = clean.length;
                            }}
                            // 0.6 is an average char-width-to-font-size ratio
                            // for proportional fonts at typical aspect ratios.
                            var textW_unrot = maxLineLen * fontSize * 0.6 * sx;
                            var textH_unrot = lines.length * fontSize * 1.25 * sy;

                            // Apply textangle rotation: if rotated near
                            // ±90°, the visual width/height swap.
                            var angle = a.textangle || 0;
                            var w, h;
                            if (Math.abs(angle) > 45) {{
                                w = textH_unrot;
                                h = textW_unrot;
                            }} else {{
                                w = textW_unrot;
                                h = textH_unrot;
                            }}

                            // Hit zone hugs the label's visual geometry.
                            // Earlier we enforced MIN_HIT_DIM = 60 with pad = 25 to
                            // make rotated-90° labels easy to hover, but those
                            // dimensions assumed a much larger label font (the
                            // pre-scaling default of 20–32 px). After scaling the
                            // visual density down to ~0.35×, the labels are now
                            // ~7-8 px tall, so a 60×60 hit zone overflows into
                            // the block area below the label and triggers the
                            // aggregation tooltip even when the cursor is on a
                            // block — producing two tooltips at once. With the
                            // hit zone sized to the label's actual extent plus
                            // a small padding buffer, only hovering the label
                            // text itself triggers the aggregation tooltip;
                            // hovering a block triggers only the block tooltip.
                            // Padding chosen as 4 px — enough tolerance for
                            // sub-pixel cursor coordinates and our em-width
                            // estimation error, but tight enough to stay clear
                            // of the block.
                            var pad = 4;

                            // Anchor offsets. For aggregation labels in
                            // this app: xanchor='center', yanchor='middle'
                            // (rotated case), so anchor is at the visual
                            // center. Default to center handling; non-
                            // center anchors get a small bias.
                            var dxLeft, dxRight, dyTop, dyBot;
                            if (a.xanchor === 'left') {{
                                dxLeft = 0; dxRight = w;
                            }} else if (a.xanchor === 'right') {{
                                dxLeft = w; dxRight = 0;
                            }} else {{ // center
                                dxLeft = w / 2; dxRight = w / 2;
                            }}
                            if (a.yanchor === 'top') {{
                                dyTop = 0; dyBot = h;
                            }} else if (a.yanchor === 'bottom') {{
                                dyTop = h; dyBot = 0;
                            }} else {{ // middle
                                dyTop = h / 2; dyBot = h / 2;
                            }}

                            zones.push({{
                                rect: {{
                                    left:   anchor_vx - dxLeft  - pad,
                                    right:  anchor_vx + dxRight + pad,
                                    top:    anchor_vy - dyTop   - pad,
                                    bottom: anchor_vy + dyBot   + pad,
                                    width:  dxLeft + dxRight + 2*pad,
                                    height: dyTop  + dyBot   + 2*pad,
                                    centerX: anchor_vx,
                                    centerY: anchor_vy
                                }},
                                text: fullText
                            }});
                        }}
                    }}
                    return zones;
                }}

                function _invalidateHitZones() {{ __ttHitZones = null; }}
                window._invalidateTtHitZones = _invalidateHitZones;

                // ── Debug visualizer ──────────────────────────────────
                // Run _showTtZones() in the console to draw every hit
                // zone as a red dashed rectangle overlaid on the page.
                // Each rect is labeled with its index. Run _hideTtZones()
                // to clear. If a zone's rectangle does NOT visually
                // overlap its aggregation label, our coordinate math is
                // off — paste the discrepancy back to me and I'll know
                // exactly where the offset is. If rectangles look right
                // but tooltip still doesn't show, the issue is elsewhere
                // (mousemove handler, hit testing, etc.).
                window._showTtZones = function() {{
                    if (window._hideTtZones) window._hideTtZones();
                    var zones = _buildHitZones();
                    zones.forEach(function(z, i) {{
                        var box = document.createElement('div');
                        box.className = '_tt_debug_box';
                        box.style.cssText = [
                            'position:fixed',
                            'border:2px dashed #ff0033',
                            'background:rgba(255,0,51,0.08)',
                            'pointer-events:none',
                            'z-index:99999',
                            'box-sizing:border-box'
                        ].join(';');
                        box.style.left   = z.rect.left   + 'px';
                        box.style.top    = z.rect.top    + 'px';
                        box.style.width  = z.rect.width  + 'px';
                        box.style.height = z.rect.height + 'px';
                        document.body.appendChild(box);

                        var label = document.createElement('div');
                        label.className = '_tt_debug_box';
                        label.style.cssText = [
                            'position:fixed',
                            'background:#ff0033',
                            'color:#fff',
                            'font-size:11px',
                            'padding:1px 4px',
                            'pointer-events:none',
                            'z-index:100000',
                            'font-family:monospace'
                        ].join(';');
                        label.style.left = z.rect.left + 'px';
                        label.style.top  = z.rect.top  + 'px';
                        label.textContent = '#' + i + ' "' + z.text.slice(0, 30) + '"';
                        document.body.appendChild(label);
                    }});
                    console.log('drew', zones.length, 'hit zones (red dashed). type _hideTtZones() to clear.');
                    return zones.length;
                }};
                window._hideTtZones = function() {{
                    document.querySelectorAll('._tt_debug_box').forEach(function(el) {{
                        el.parentNode && el.parentNode.removeChild(el);
                    }});
                }};

                // Invalidate on viewport / pan-zoom changes.
                window.addEventListener('scroll', _invalidateHitZones, {{ passive: true }});
                window.addEventListener('resize', _invalidateHitZones);
                window.addEventListener('wheel',  _invalidateHitZones, {{ passive: true }});

                document.addEventListener('mousemove', function(e) {{
                    var cx = e.clientX, cy = e.clientY;
                    var td = _getTtDiv();

                    // Lazy rebuild after invalidation. The build runs at
                    // a moment when no plotly hover is in progress
                    // (cursor is either far from any label or only just
                    // entering one), so SVG state is unmutated and the
                    // bboxes captured are the labels' true positions.
                    if (__ttHitZones === null) {{
                        __ttHitZones = _buildHitZones();
                    }}

                    // Hit test against cache — NO DOM queries here.
                    // Cache snapshot survives plotly's runtime SVG
                    // mutations during hover, so the tooltip stays
                    // visible as long as the cursor is over the cached
                    // hit zone.
                    var found = null;
                    var foundRect = null;
                    var bestArea = Infinity;
                    for (var i = 0; i < __ttHitZones.length; i++) {{
                        var z = __ttHitZones[i];
                        var r = z.rect;
                        if (cx < r.left || cx > r.right) continue;
                        if (cy < r.top  || cy > r.bottom) continue;
                        var area = r.width * r.height;
                        if (area < bestArea) {{
                            bestArea = area;
                            found = z.text;
                            foundRect = r;
                        }}
                    }}

                    if (found && foundRect) {{
                        // Pin tooltip to the LABEL's bbox — placement
                        // is stable as long as the cursor is anywhere
                        // over the cached hit zone.
                        td.textContent = found;
                        td.style.display = 'block';
                        // Mark a flag the plotly_hover handler will read
                        // to suppress plotly's native hover label entirely.
                        // The simple "call unhover() in mousemove" approach
                        // races against plotly's own hover renderer: by the
                        // time mousemove fires plotly may have already
                        // painted the band/block hover label, and the
                        // unhover call only catches the next frame, leaving
                        // a brief overlap visible. The flag-driven approach
                        // intercepts at the plotly_hover event itself, BEFORE
                        // any rendering happens, so the band/block tooltip
                        // never appears at all when the cursor is also over
                        // an enclosure label.
                        window._enclosureTtActive = true;
                        var tw = td.offsetWidth, th = td.offsetHeight;

                        // Also issue a one-shot unhover for the case where
                        // a plotly hover label was already displayed before
                        // the cursor entered our hit zone (catches the
                        // initial-state transition that plotly_hover
                        // suppression alone would miss).
                        try {{
                            var _gdH = (typeof getPlotlyGd === 'function') ? getPlotlyGd() : null;
                            if (_gdH && typeof Plotly !== 'undefined' && Plotly.Fx && Plotly.Fx.unhover) {{
                                Plotly.Fx.unhover(_gdH);
                            }}
                        }} catch (e) {{}}

                        // Default: 10px right of bbox, vertically
                        // centered with the label.
                        var lx = foundRect.right + 10;
                        var ly = foundRect.top + (foundRect.height - th) / 2;

                        // Edge fallbacks (right → left → above → below)
                        if (lx + tw > window.innerWidth - 8) {{
                            lx = foundRect.left - tw - 10;
                            if (lx < 8) {{
                                lx = Math.max(8, Math.min(foundRect.left,
                                    window.innerWidth - tw - 8));
                                ly = foundRect.top - th - 10;
                                if (ly < 8) ly = foundRect.bottom + 10;
                            }}
                        }}
                        if (ly < 8) ly = 8;
                        if (ly + th > window.innerHeight - 8) {{
                            ly = window.innerHeight - th - 8;
                        }}

                        td.style.left = lx + 'px';
                        td.style.top  = ly + 'px';
                    }} else {{
                        td.style.display = 'none';
                        window._enclosureTtActive = false;
                    }}
                }}, {{ passive: true }});

                // ── Defensive cleanup: hide tooltip whenever the cursor
                //    leaves the viewport, the tab loses focus/visibility,
                //    or the Plotly canvas relayouts/redraws. Without these,
                //    the floating div can be left pinned to the last mouse
                //    position when the user alt-tabs, opens a modal, or
                //    when Plotly mutates the annotation layer underneath.
                function _hideLblTT() {{
                    if (_ttDiv) _ttDiv.style.display = 'none';
                    window._enclosureTtActive = false;
                }}
                document.documentElement.addEventListener('mouseleave', _hideLblTT);
                window.addEventListener('blur', _hideLblTT);
                document.addEventListener('visibilitychange', function() {{
                    if (document.hidden) _hideLblTT();
                }});
                window._hideLblTT = _hideLblTT;   // exposed for plotly_afterplot hookup
            }})();
            function attachTruncatedLabelTooltips(gd) {{
                if (window._attachTruncatedLabelTooltips) window._attachTruncatedLabelTooltips(gd);
            }}

            // ── DPR normalisation ─────────────────────────────────────────────
            // Plotly font sizes and line widths are absolute CSS px.
            // At DPR > 1 (Windows 125% scaling), CSS px render physically larger.
            // Block/enclosure sizes are data units and DPR-compensate automatically
            // via CSS transform (canvasZoom × DPR = constant physical size).
            // This function scales all absolute-px values by 1/DPR so they match
            // their physical size across all browsers.
            //
            // Why JS and not Python: Python doesn't know DPR at first render time
            // (store-graph-size isn't populated before the first rebuild_base fires).
            // JS always has the correct DPR with no timing issue.
            //
            // ── Cross-browser diagnostic ──────────────────────────────────
            // Run `_browserDiag()` in the console of each browser and paste
            // the output. Compares all the numbers that influence layout +
            // rendering: viewport, DPR, plotly's internal layout dimensions,
            // CSS computed font sizes, SVG sample geometry, etc. The first
            // value that differs between browsers IS the root cause.
            window._browserDiag = function() {{
                var d = {{}};
                d.userAgent = navigator.userAgent;
                d.platform = navigator.platform;
                d.devicePixelRatio = window.devicePixelRatio;
                d.innerWidth  = window.innerWidth;
                d.innerHeight = window.innerHeight;
                if (typeof screen !== 'undefined') {{
                    d.screenWidth  = screen.width;
                    d.screenHeight = screen.height;
                    d.screenAvailWidth = screen.availWidth;
                }}
                d.documentClientWidth  = document.documentElement.clientWidth;
                d.documentClientHeight = document.documentElement.clientHeight;
                var graphWrap = document.getElementById('graph-wrap');
                if (graphWrap) {{
                    var r = graphWrap.getBoundingClientRect();
                    d.graphWrap_rect = {{w: r.width, h: r.height, l: r.left, t: r.top}};
                    d.graphWrap_offsetW = graphWrap.offsetWidth;
                    d.graphWrap_offsetH = graphWrap.offsetHeight;
                }}
                var gd = document.getElementById('graph');
                if (gd) {{
                    var rg = gd.getBoundingClientRect();
                    d.graph_rect = {{w: rg.width, h: rg.height}};
                    d.graph_offsetW = gd.offsetWidth;
                    d.graph_offsetH = gd.offsetHeight;
                    if (gd._fullLayout) {{
                        d.fullLayout_width  = gd._fullLayout.width;
                        d.fullLayout_height = gd._fullLayout.height;
                        d.fullLayout_font_size = gd._fullLayout.font && gd._fullLayout.font.size;
                        d.fullLayout_meta_base_sizes = !!(gd._fullLayout.meta && gd._fullLayout.meta.base_sizes);
                        if (gd._fullLayout.meta && gd._fullLayout.meta.base_sizes) {{
                            d.base_sizes_layout_font  = gd._fullLayout.meta.base_sizes.layout_font_size;
                            d.base_sizes_n_traces     = (gd._fullLayout.meta.base_sizes.traces || []).length;
                            d.base_sizes_first_trace_tf = (gd._fullLayout.meta.base_sizes.traces || [])[0];
                            d.base_sizes_first_anno_fs  = (gd._fullLayout.meta.base_sizes.annos  || [])[0];
                        }}
                        // Sample first 2 trace text sizes as currently in fullLayout.
                        var traces = gd._fullLayout.data || [];
                        d.live_trace0_textfont = traces[0] && traces[0].textfont && traces[0].textfont.size;
                        d.live_trace0_line_width = traces[0] && traces[0].line && traces[0].line.width;
                        // Sample first annotation font.
                        var annos = gd._fullLayout.annotations || [];
                        d.live_anno0_font_size = annos[0] && annos[0].font && annos[0].font.size;
                        d.live_n_annotations = annos.length;
                    }}
                }}
                // Sample SVG <text> rendered size — this is what user actually sees.
                var firstSvgText = document.querySelector('.js-plotly-plot text');
                if (firstSvgText) {{
                    var cs = window.getComputedStyle(firstSvgText);
                    d.first_svg_text_computed_fontSize = cs.fontSize;
                    d.first_svg_text_attr_fontSize = firstSvgText.getAttribute('font-size');
                    var br = firstSvgText.getBoundingClientRect();
                    d.first_svg_text_bbox = {{w: br.width, h: br.height}};
                }}
                // Sample first SVG shape stroke-width.
                var firstShape = document.querySelector('.js-plotly-plot .shapelayer path');
                if (firstShape) {{
                    var cs2 = window.getComputedStyle(firstShape);
                    d.first_shape_strokeWidth = cs2.strokeWidth;
                    d.first_shape_attr_strokeWidth = firstShape.getAttribute('stroke-width');
                }}
                // Canvas zoom / transform.
                var cc = document.getElementById('canvas-content');
                if (cc) {{
                    d.canvas_content_transform = cc.style.transform;
                }}
                d.css_var_canvas_zoom = getComputedStyle(document.documentElement).getPropertyValue('--canvas-zoom');
                console.log(JSON.stringify(d, null, 2));
                return d;
            }};

            function applyDPRScaling(gd) {{
                // ── Disabled: this function previously multiplied font sizes,
                // border widths, and line widths by 1/DPR (or by viewport
                // ratio in a later iteration) to "normalize" visual size
                // across devices. Both approaches turned out to be
                // double-corrections — browsers already map CSS px to
                // physical px using system DPI, and the chart container
                // already scales with viewport via CSS. Multiplying again
                // produced inconsistent visuals across browsers, especially
                // in edge cases like multi-monitor setups with mixed DPI
                // (an external display attached to a high-DPI laptop) where
                // some browsers (Lenovo's Chromium fork in particular)
                // report wildly inflated innerWidth values.
                //
                // Removing the scaling makes rendering deterministic on the
                // common path: cfg.block_text_size = 7 means Plotly draws
                // SVG text at 7 CSS px on every browser. Reviewers running
                // standard browsers (Chrome, Firefox, Edge, Safari) at 100%
                // zoom on a single monitor — the overwhelming majority case
                // — see consistent, reasonable rendering. Edge cases like
                // multi-monitor mixed DPI on niche browsers will still
                // exhibit browser-level quirks but they're not the target
                // audience.
                return;
                var meta = gd._fullLayout && gd._fullLayout.meta;
                var bs   = meta && meta.base_sizes;

                // ── Trace text fonts (block labels) ─────────────────────────
                // base_sizes.traces has DPR=1 reference values from Python.
                // Safe to apply repeatedly — always scales from the reference.
                if (bs && bs.traces) {{
                    for (var i = 0; i < bs.traces.length && i < gd.data.length; i++) {{
                        var bt = bs.traces[i], tr = gd.data[i];
                        if (!bt || !tr) continue;
                        try {{ if (bt.tf != null && tr.textfont) tr.textfont.size = bt.tf * sc; }} catch(e) {{}}
                    }}
                }}

                // ── Shape line widths ─────────────────────────────────────────
                // base_sizes.shapes covers shapes from rebuild_base.
                // Extra shapes (highlight overlays added by render_with_highlight) don't
                // have a base — scale them using current-value tracking to avoid drift.
                if (bs && bs.shapes && gd._fullLayout.shapes) {{
                    var nBase = bs.shapes.length;
                    for (var i = 0; i < nBase && i < gd._fullLayout.shapes.length; i++) {{
                        var bsh = bs.shapes[i], sh = gd._fullLayout.shapes[i];
                        if (!bsh || bsh.lw == null || !sh || !sh.line) continue;
                        try {{ sh.line.width = bsh.lw * sc; }} catch(e) {{}}
                    }}
                    // Extra (highlight overlay) shapes: scale current value once
                    for (var i = nBase; i < gd._fullLayout.shapes.length; i++) {{
                        var sh = gd._fullLayout.shapes[i];
                        if (!sh || !sh.line) continue;
                        try {{
                            if (!sh._dprScaled) {{
                                sh._dprScaled = true;
                                sh.line.width = (sh.line.width || 0) * sc;
                            }}
                        }} catch(e) {{}}
                    }}
                }}

                // ── Trace line widths ────────────────────────────────────────
                // Some traces are modified by render_with_highlight (selected block/band
                // gets a different line width). We handle both highlighted and base widths
                // by detecting when the current value differs from the base_sizes reference.
                if (bs && bs.traces) {{
                    for (var i = 0; i < bs.traces.length && i < gd.data.length; i++) {{
                        var bt = bs.traces[i], tr = gd.data[i];
                        if (!bt || !tr || !tr.line || bt.lw == null) continue;
                        try {{
                            var cur = tr.line.width;
                            if (Math.abs(cur - bt.lw) < 0.1) {{
                                // Not highlighted: scale from base
                                tr.line.width = bt.lw * sc;
                            }} else {{
                                // Highlighted (or otherwise changed): scale current, track to avoid drift
                                if (tr._dprLwBase == null || Math.abs(cur - tr._dprLwBase * sc) > 0.1) {{
                                    tr._dprLwBase = cur;
                                }}
                                tr.line.width = tr._dprLwBase * sc;
                            }}
                        }} catch(e) {{}}
                    }}
                }}

                // ── Annotation font sizes ─────────────────────────────────────
                // Annotations change between renders (label styles, visibility).
                // Self-correcting: detect if Python changed font size, update base, re-scale.
                (gd._fullLayout.annotations || []).forEach(function(ann, i) {{
                    if (!ann || !ann.font || ann.font.size == null) return;
                    var cur = ann.font.size;
                    try {{
                        if (ann._dprBase == null) {{
                            ann._dprBase = cur;
                        }} else {{
                            var expected = ann._dprBase * sc;
                            if (Math.abs(cur - expected) > 0.3) {{
                                ann._dprBase = cur;  // Python changed font size, update base
                            }}
                        }}
                        ann.font.size = ann._dprBase * sc;
                    }} catch(e) {{}}
                }});

                // ── Layout font + hoverlabel ──────────────────────────────────
                // Layout font: still scaled by 1/DPR (legacy behavior, used for
                // axis titles, ticklabels, etc. which don't have a canvasZoom
                // counter-scale path).
                //
                // Hoverlabel font: ownership transferred to applyCanvasTransform.
                // applyCanvasTransform sets hoverlabel.font.size with a formula
                // that already accounts for both viewport width AND canvasZoom,
                // matching the HTML tooltip's vw-based sizing exactly. We must
                // NOT overwrite it here — that previously caused block/band
                // tooltips to drift away from enclosure/aggregation tooltip
                // sizing whenever DPR > 1 (Windows 125%/150% scaling).
                var lfsBase = bs && bs.layout_font_size;
                try {{ if (lfsBase && gd._fullLayout.font) gd._fullLayout.font.size = lfsBase * sc; }} catch(e) {{}}
            }}


            (function tryAttachAfterPlot() {{
                var gd = getPlotlyGd();
                if (!gd || typeof gd.on !== 'function') {{
                    setTimeout(tryAttachAfterPlot, 200);
                    return;
                }}
                gd.on('plotly_afterplot', function() {{
                    if (typeof Plotly !== 'undefined') Plotly.Plots.resize(gd);
                    applyCanvasTransform();
                    applyContainFit(gd);
                    applyDPRScaling(gd);
                    attachTruncatedLabelTooltips(gd);
                    // ── Do NOT call _hideLblTT() here ───────────────────────
                    // Plotly fires plotly_afterplot during its OWN hover label
                    // rendering (on hover-in / hover-out of any trace). Calling
                    // _hideLblTT here would forcibly hide our aggregation
                    // tooltip every time the user moves between traces,
                    // creating a flicker (visible → hidden by afterplot →
                    // visible again on next mousemove → hidden by afterplot…).
                    // Tooltip is hidden when the cursor leaves the cached
                    // hit zone (handled in mousemove); the cache itself is
                    // invalidated below so positions stay correct after
                    // genuine redraws.
                    if (window._invalidateTtHitZones) window._invalidateTtHitZones();

                    // ── Status badge: clear pending_render on render done ──
                    // afterplot fires for genuine redraws AND for hover label
                    // rendering. Use the `expecting_afterplot` gate (set in
                    // the figure-change clientside callback) to ignore
                    // hover-induced afterplot. Only the first afterplot
                    // AFTER a figure prop change consumes the gate and
                    // clears pending_render.
                    if (window._birdcage_expecting_afterplot) {{
                        window._birdcage_expecting_afterplot = false;
                        if (window._birdcage_status_state &&
                            window._birdcage_status_state.pending_render) {{
                            window._birdcage_status_state.pending_render = false;
                            if (window._birdcage_pending_render_timer) {{
                                clearTimeout(window._birdcage_pending_render_timer);
                                window._birdcage_pending_render_timer = null;
                            }}
                            if (window._birdcage_update_status_badge) {{
                                window._birdcage_update_status_badge();
                            }}
                        }}
                    }}
                }});

                // ── Single-tooltip enforcement ────────────────────────────────
                // When the cursor is over an enclosure-name hit zone, our
                // mousemove handler sets window._enclosureTtActive = true.
                // Plotly may simultaneously detect a band or block under the
                // same cursor position and try to render its native hover
                // label — overlapping our tooltip (the user reported this
                // exact case: "Merge" band hover stacking on top of a
                // "Food Services and Drinking Places" enclosure tooltip).
                //
                // plotly_hover fires BEFORE the hover label is rendered, so
                // calling Fx.unhover here cancels rendering before it
                // happens. Net effect: only one tooltip on screen at any
                // moment — the enclosure one wins when the cursor is on a
                // label, plotly's wins everywhere else.
                gd.on('plotly_hover', function() {{
                    if (window._enclosureTtActive) {{
                        try {{
                            if (typeof Plotly !== 'undefined' && Plotly.Fx && Plotly.Fx.unhover) {{
                                Plotly.Fx.unhover(gd);
                            }}
                        }} catch (e) {{}}
                    }}
                }});
            }})();

            // ── Hide Plotly's native annotation hover tooltip ─────────────
            // The CSS suppression above (.annotation-hovertext etc.) covers
            // older Plotly versions, but recent Plotly versions (Dash 2.x+
            // bundles) render annotation hover labels through Fx.loneHover
            // into the same .hoverlayer that trace tooltips use, with the
            // SAME `g.hovertext` class — making CSS-based discrimination
            // impossible by class alone.
            //
            // The drift symptom (tooltip pinned to annotation's data
            // anchor instead of the cursor) is caused by this leaked
            // native tooltip. The custom HTML overlay (#_lbl_tt) is
            // already there and follows the cursor correctly, but the
            // native tooltip renders ON TOP of the canvas SVG and looks
            // identical because we styled them the same way.
            //
            // Discrimination strategy: by CONTENT. Our annotation
            // hovertexts are the FULL aggregation names ("RV (Recreational
            // Vehicle) Parks and Recreational Camps", etc.). Trace hover
            // labels show block/band names — never these exact strings.
            // We watch for any new <g class="hovertext"> being added to
            // the DOM, and if its text matches an annotation.hovertext,
            // we hide it. Trace tooltips are untouched.
            //
            // MutationObserver fires synchronously between mutation and
            // paint, so there's no visible flicker of the native tooltip
            // before we hide it.
            (function setupNativeAnnoTooltipHider() {{
                function normalizeText(s) {{
                    return String(s || '').replace(/\\s+/g, ' ').trim();
                }}
                function matchesAnyAnnoHovertext(content) {{
                    if (!content) return false;
                    var charts = document.querySelectorAll('.js-plotly-plot');
                    for (var ci = 0; ci < charts.length; ci++) {{
                        var fl = charts[ci]._fullLayout;
                        if (!fl) continue;
                        var annos = fl.annotations || [];
                        for (var i = 0; i < annos.length; i++) {{
                            var a = annos[i];
                            if (!a || !a.hovertext) continue;
                            if (normalizeText(a.hovertext) === content) return true;
                        }}
                    }}
                    return false;
                }}
                function isNativeAnnoTooltip(node) {{
                    if (!node || node.nodeType !== 1) return false;
                    if (!node.classList || !node.classList.contains('hovertext')) return false;
                    var t = node.querySelector('text');
                    if (!t) return false;
                    return matchesAnyAnnoHovertext(normalizeText(t.textContent));
                }}
                function checkAndHide(node) {{
                    if (isNativeAnnoTooltip(node)) {{
                        node.style.display = 'none';
                        node.setAttribute('data-anno-native-hidden', '1');
                    }}
                }}
                var observer = new MutationObserver(function(muts) {{
                    for (var mi = 0; mi < muts.length; mi++) {{
                        var m = muts[mi];
                        if (m.type !== 'childList') continue;
                        for (var ni = 0; ni < m.addedNodes.length; ni++) {{
                            var node = m.addedNodes[ni];
                            checkAndHide(node);
                            // Defensive: also check descendants in case
                            // the .hovertext was wrapped in a parent group
                            // that was added in one mutation.
                            if (node.nodeType === 1 && node.querySelectorAll) {{
                                var inner = node.querySelectorAll('g.hovertext');
                                for (var ii = 0; ii < inner.length; ii++) {{
                                    checkAndHide(inner[ii]);
                                }}
                            }}
                        }}
                    }}
                }});
                // Observe document.body with subtree:true — survives chart
                // DOM replacement (Dash purge+newPlot path) without needing
                // to re-attach the observer.
                function start() {{
                    if (!document.body) {{ setTimeout(start, 100); return; }}
                    observer.observe(document.body, {{ childList: true, subtree: true }});
                    // Also sweep currently-visible hovertexts (in case
                    // observer attached after one was already inserted).
                    document.querySelectorAll('g.hovertext').forEach(checkAndHide);
                }}
                start();
                window._nativeAnnoTooltipObserver = observer;   // exposed for debugging
            }})();

            // ── Force category-combo dropdown menu to fit ALL options
            //    without any scrollbar, regardless of option count.
            //    dcc.Dropdown wraps react-virtualized-select, which renders
            //    options inside a Grid that has an inline-style max-height
            //    (~200px by default) and a scrolling inner ScrollContent.
            //    CSS !important alone can't beat the inline-style on the
            //    inner Grid, so we use a MutationObserver: every time the
            //    menu DOM appears or changes, we measure the natural
            //    summed height of all option rows and write that height
            //    onto the Grid (and its parent), eliminating the scroll.
            (function fitCategoryMenu() {{
                var combo = null;
                function setup() {{
                    combo = document.getElementById('category-combo');
                    if (!combo) {{ setTimeout(setup, 200); return; }}
                    var observer = new MutationObserver(fixMenu);
                    observer.observe(combo, {{
                        childList: true, subtree: true, attributes: true,
                        attributeFilter: ['style', 'class']
                    }});
                    fixMenu();
                }}
                function fixMenu() {{
                    if (!combo) return;
                    var menuOuter = combo.querySelector('.Select-menu-outer');
                    if (!menuOuter) return;
                    // Measure the total natural height: sum of every visible
                    // option row's height (use scrollHeight as the floor —
                    // react-virtualized renders only the visible window of
                    // rows, but its inner content's scrollHeight equals the
                    // total height of ALL rows).
                    var grids = menuOuter.querySelectorAll('[class*="Grid"], .ReactVirtualized__Grid, .ReactVirtualized__Grid__innerScrollContainer');
                    var maxScroll = 0;
                    for (var i = 0; i < grids.length; i++) {{
                        var g = grids[i];
                        if (g.scrollHeight && g.scrollHeight > maxScroll) maxScroll = g.scrollHeight;
                    }}
                    // Fallback if Grid not found yet — use option row × count.
                    if (maxScroll === 0) {{
                        var rows = menuOuter.querySelectorAll('.Select-option, .VirtualizedSelectOption');
                        if (rows.length > 0) {{
                            maxScroll = rows[0].offsetHeight * rows.length;
                        }}
                    }}
                    if (maxScroll === 0) return;
                    var hPx = maxScroll + 'px';
                    // Apply to every Grid-like container so neither the
                    // outer max-height nor the inner one cap the menu.
                    for (var j = 0; j < grids.length; j++) {{
                        grids[j].style.setProperty('max-height', 'none', 'important');
                        grids[j].style.setProperty('height', hPx, 'important');
                        grids[j].style.setProperty('overflow', 'visible', 'important');
                    }}
                    menuOuter.style.setProperty('max-height', 'none', 'important');
                    menuOuter.style.setProperty('overflow', 'visible', 'important');
                    var menu = menuOuter.querySelector('.Select-menu');
                    if (menu) {{
                        menu.style.setProperty('max-height', 'none', 'important');
                        menu.style.setProperty('overflow', 'visible', 'important');
                        menu.style.setProperty('height', hPx, 'important');
                    }}
                }}
                setup();
            }})();

            console.log('Canvas controls initialized (CSS transform zoom/pan)');
            // Apply initial UI scale so right panel = 1/6 of viewport on first load.
            // Temporarily disable the transform transition — without this, Safari and
            // Firefox animate FROM "no transform" (scale 1) TO the computed scale on
            // first paint, producing a visible ~50ms shrink. We restore transitions
            // on the next frame so subsequent user-driven scales still animate.
            var _savedTransitions = [];
            ['ui-top-wrap', 'right-panel-inner'].forEach(function(id) {{
                var el = document.getElementById(id);
                if (!el) return;
                _savedTransitions.push([el, el.style.transition]);
                el.style.transition = 'none';
            }});
            applyUITransform();
            requestAnimationFrame(function() {{
                requestAnimationFrame(function() {{
                    _savedTransitions.forEach(function(pair) {{ pair[0].style.transition = pair[1]; }});
                }});
            }});
            // One-shot delayed call: catches the very first render before plotly_afterplot
            // listener is confirmed to be attached.
            // Initial call: run after layout is settled so graph-wrap.offsetWidth is accurate
            setTimeout(applyUITransform, 50);
            // Second call after Plotly has had time to render and resize graph-wrap
            setTimeout(applyUITransform, 600);
        }}

        // Enhance number inputs with custom +/- buttons
        function enhanceNumberInputs() {{
            var rightPanel = document.getElementById('right-panel-inner');
            var downloadPanel = document.getElementById('download-panel');
            if (!rightPanel && !downloadPanel) {{
                setTimeout(enhanceNumberInputs, 200);
                return;
            }}

            var containers = [];
            if (rightPanel) containers.push(rightPanel);
            if (downloadPanel) containers.push(downloadPanel);

            containers.forEach(function(container) {{
                var inputs = container.querySelectorAll('input[type="number"]');
                inputs.forEach(function(input) {{
                // Skip if already enhanced
                if (input.dataset.enhanced === 'true') return;
                input.dataset.enhanced = 'true';

                // Get step value
                var step = parseFloat(input.step) || 1;
                var min = input.min !== '' ? parseFloat(input.min) : -Infinity;
                var max = input.max !== '' ? parseFloat(input.max) : Infinity;

                // Create wrapper
                // NOTE: some number inputs (e.g., rotation controls) are intended to be full-width.
                // Using inline-flex causes the wrapper to shrink-to-fit and prevents width:100% inputs from stretching.
                var wrapper = document.createElement('div');
                var isFullWidth = (input.classList && input.classList.contains('fullwidth-number'));
                if (isFullWidth) {{
                    wrapper.style.cssText = 'display:flex; align-items:stretch; position:relative; width:100%; min-width:0; box-sizing:border-box;';
                }} else {{
                    wrapper.style.cssText = 'display:inline-flex; align-items:stretch; position:relative;';
                }}

                // Style the input to hide native spinner
                input.style.cssText += '; -webkit-appearance:textfield; -moz-appearance:textfield; appearance:textfield; padding-right:32px !important;';

                // Insert wrapper
                input.parentNode.insertBefore(wrapper, input);
                wrapper.appendChild(input);

                // Create button container
                var btnContainer = document.createElement('div');
                btnContainer.style.cssText = 'position:absolute; right:2px; top:2px; bottom:2px; width:28px; display:flex; flex-direction:column; border-left:1px solid #CCCCCC; background:#F8F8F8; border-radius:0 4px 4px 0; overflow:hidden;';
                wrapper.appendChild(btnContainer);

                // Create up button
                var upBtn = document.createElement('button');
                upBtn.type = 'button';
                upBtn.className = 'spinner-btn';
                upBtn.innerHTML = '▲';
                upBtn.style.cssText = 'flex:1; border:none; background:transparent; cursor:pointer; font-size:10px; color:#555; padding:0; line-height:1; display:flex; align-items:center; justify-content:center; height:auto; min-height:0;';
                upBtn.onmouseover = function() {{ this.style.background='#E8E8E8'; }};
                upBtn.onmouseout = function() {{ this.style.background='transparent'; }};
                upBtn.onclick = function(e) {{
                    e.preventDefault();
                    e.stopPropagation();
                    var val = parseFloat(input.value) || 0;
                    var newVal = Math.min(max, val + step);
                    // Round to avoid floating point issues
                    newVal = Math.round(newVal * 1000000) / 1000000;
                    // Use native setter to trigger React's change detection
                    var nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
                    nativeInputValueSetter.call(input, newVal);
                    // Dispatch events for React
                    var inputEvent = new Event('input', {{ bubbles: true, cancelable: true }});
                    input.dispatchEvent(inputEvent);
                }};
                btnContainer.appendChild(upBtn);

                // Create separator
                var sep = document.createElement('div');
                sep.style.cssText = 'height:1px; background:#DDDDDD; flex-shrink:0;';
                btnContainer.appendChild(sep);

                // Create down button
                var downBtn = document.createElement('button');
                downBtn.type = 'button';
                downBtn.className = 'spinner-btn';
                downBtn.innerHTML = '▼';
                downBtn.style.cssText = 'flex:1; border:none; background:transparent; cursor:pointer; font-size:10px; color:#555; padding:0; line-height:1; display:flex; align-items:center; justify-content:center; height:auto; min-height:0;';
                downBtn.onmouseover = function() {{ this.style.background='#E8E8E8'; }};
                downBtn.onmouseout = function() {{ this.style.background='transparent'; }};
                downBtn.onclick = function(e) {{
                    e.preventDefault();
                    e.stopPropagation();
                    var val = parseFloat(input.value) || 0;
                    var newVal = Math.max(min, val - step);
                    // Round to avoid floating point issues
                    newVal = Math.round(newVal * 1000000) / 1000000;
                    // Use native setter to trigger React's change detection
                    var nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
                    nativeInputValueSetter.call(input, newVal);
                    // Dispatch events for React
                    var inputEvent = new Event('input', {{ bubbles: true, cancelable: true }});
                    input.dispatchEvent(inputEvent);
                }};
                btnContainer.appendChild(downBtn);
                }});
            }});

            // Re-run periodically to catch dynamically added inputs
            setTimeout(enhanceNumberInputs, 2000);

            // Re-enhance download panel inputs whenever the Download button is clicked
            (function() {{
                var btnDl = document.getElementById('btn-download');
                if (!btnDl) return;
                btnDl.addEventListener('click', function() {{
                    setTimeout(function() {{
                        var dlPanel = document.getElementById('download-panel');
                        if (!dlPanel) return;
                        dlPanel.querySelectorAll('input[type="number"]').forEach(function(inp) {{
                            var par = inp.parentNode;
                            // Unwrap previous wrapper if present
                            if (par && par.style && par.style.position === 'relative' && par !== dlPanel) {{
                                par.parentNode.insertBefore(inp, par);
                                par.parentNode.removeChild(par);
                            }}
                            inp.dataset.enhanced = 'false';
                        }});
                        enhanceNumberInputs();
                    }}, 50);
                }});
            }})();
        }}

        // =====================================================
        // Popover dragging functionality
        // =====================================================
        function initPopoverDrag() {{
            var popovers = ['file-details-popover', 'download-panel', 'ordering-panel', 'guide-popover'];

            popovers.forEach(function(popoverId) {{
                var popover = document.getElementById(popoverId);
                if (!popover || popover.dataset.dragInitialized === 'true') return;
                popover.dataset.dragInitialized = 'true';

                var isDragging = false;
                var startX, startY, startLeft, startTop;

                popover.addEventListener('mousedown', function(e) {{
                    // Only start drag if clicking on the popover itself or its header area (not controls)
                    if (e.target.tagName === 'INPUT' || e.target.tagName === 'SELECT' || 
                        e.target.tagName === 'BUTTON' || e.target.closest('.Select') ||
                        e.target.closest('button') || e.target.id.includes('close')) {{
                        return;
                    }}

                    isDragging = true;
                    startX = e.clientX;
                    startY = e.clientY;

                    var style = window.getComputedStyle(popover);
                    var left = parseFloat(style.left) || 0;
                    var top = parseFloat(style.top) || 0;

                    // If using percentage-based positioning, convert to pixels
                    if (style.top.includes('%')) {{
                        var parentRect = popover.offsetParent ? popover.offsetParent.getBoundingClientRect() : document.body.getBoundingClientRect();
                        top = popover.offsetTop;
                    }}

                    startLeft = left;
                    startTop = top;

                    popover.style.cursor = 'grabbing';
                    e.preventDefault();
                }});

                document.addEventListener('mousemove', function(e) {{
                    if (!isDragging) return;

                    var dx = e.clientX - startX;
                    var dy = e.clientY - startY;

                    // Compensate for scaled ancestors. The popover lives inside
                    // #ui-top-wrap which is transform:scale(uiScale). mousemove
                    // deltas are in viewport pixels, but style.left/top are in
                    // layout pixels of the (pre-transform) parent. Dividing by
                    // the effective scale keeps the popover locked to the cursor
                    // regardless of UI zoom.
                    var effScale = 1.0;
                    try {{
                        var ancestor = popover.parentElement;
                        while (ancestor && ancestor !== document.body) {{
                            var tr = window.getComputedStyle(ancestor).transform;
                            if (tr && tr !== 'none') {{
                                // matrix(a, b, c, d, tx, ty) — a is the x-scale
                                var m = tr.match(/matrix\\(([^)]+)\\)/);
                                if (m) {{
                                    var parts = m[1].split(',').map(parseFloat);
                                    if (parts.length >= 4 && isFinite(parts[0]) && parts[0] > 0) {{
                                        effScale *= parts[0];
                                    }}
                                }}
                            }}
                            ancestor = ancestor.parentElement;
                        }}
                    }} catch (err) {{ /* fall back to scale 1 */ }}
                    if (!(effScale > 0)) effScale = 1.0;

                    popover.style.left = (startLeft + dx / effScale) + 'px';
                    popover.style.top = (startTop + dy / effScale) + 'px';
                    popover.style.position = 'absolute';
                }});

                document.addEventListener('mouseup', function() {{
                    if (isDragging) {{
                        isDragging = false;
                        popover.style.cursor = 'move';
                        // Sync the just-dragged position into the shared
                        // `store-popover-positions` Dash store so that closing
                        // and reopening the popover restores this location
                        // (instead of snapping back to the default anchor).
                        try {{
                            if (!window._popoverPositions) window._popoverPositions = {{}};
                            window._popoverPositions[popoverId] = {{
                                left: popover.style.left || '',
                                top:  popover.style.top  || '',
                            }};
                            var trig = document.getElementById('popover-drag-trigger');
                            if (trig) {{
                                trig.dataset.payload = JSON.stringify(window._popoverPositions);
                                trig.click();
                            }}
                        }} catch (err) {{ /* swallow — non-critical */ }}
                    }}
                }});
            }});

            // Re-run periodically to catch dynamically shown popovers
            setTimeout(initPopoverDrag, 1000);
        }}

        // ── Spacer-based right-panel alignment ───────────────────────
        // right-panel has top:0 always. A spacer at the top of the panel has
        // layout height = topBar.offsetHeight. Both are scaled by the same uiScale,
        // so their visual heights always match — no JS top-sync needed.
        var _lastSyncH = -1;
        function syncRightPanelTop() {{
            var topBar = document.getElementById('ui-top-wrap');
            var spacer  = document.getElementById('right-panel-topspacer');
            var graphWrap = document.getElementById('graph-wrap');
            if (!topBar) return;
            // Compute uiScale locally — matches the formula in initCanvasControls.
            // Previously this referenced an outer-scope `uiScale` which is
            // a closure variable inside another function; this function runs
            // at global scope (called from DOMContentLoaded / ResizeObserver),
            // so the closure binding wasn't visible here → ReferenceError on
            // every event the observer fired. Computing locally with the same
            // formula keeps the value in lockstep at zero coupling cost.
            var _NATURAL_PANEL_W = 920;
            var _uiScale = Math.min(2.0, Math.max(0.15,
                              window.innerWidth / (6 * _NATURAL_PANEL_W)));
            // offsetHeight = layout height, unaffected by CSS transform.
            // Visual top-bar height = offsetHeight × uiScale.
            var layoutH = topBar.offsetHeight;
            if (layoutH <= 0) return;
            var h = Math.round(layoutH * _uiScale);   // visual bottom of top-bar
            if (h === _lastSyncH) return;
            _lastSyncH = h;
            var availH = window.innerHeight - h;
            // Keep spacer layout height in sync with top-bar layout height.
            // (Normally constant at 84px; grows if upload-status appears.)
            if (spacer) {{
                spacer.style.height    = layoutH + 'px';
                spacer.style.minHeight = layoutH + 'px';
            }}
            // Graph-wrap has no transform — layout = visual.
            if (graphWrap) {{
                graphWrap.style.top    = h + 'px';
                graphWrap.style.height = availH + 'px';
            }}
            window.dispatchEvent(new Event('resize'));
        }}

        // Re-sync whenever the top-bar itself changes height (e.g. upload-status appears)
        (function() {{
            var topBar = document.getElementById('ui-top-wrap');
            if (topBar && window.ResizeObserver) {{
                new ResizeObserver(function() {{ syncRightPanelTop(); }}).observe(topBar);
            }}
            // Polling fallback: retry 20 times × 150 ms = 3 s after page load
            // to catch cases where Dash renders the top-bar later than DOMContentLoaded.
            var _syncPoll = 0;
            var _syncTimer = setInterval(function() {{
                syncRightPanelTop();
                if (++_syncPoll >= 20) clearInterval(_syncTimer);
            }}, 150);
        }})();


        // Start initialization
        if (document.readyState === 'loading') {{
            document.addEventListener('DOMContentLoaded', initCanvasControls);
            document.addEventListener('DOMContentLoaded', function() {{ setTimeout(enhanceNumberInputs, 500); }});
            document.addEventListener('DOMContentLoaded', function() {{ setTimeout(initPopoverDrag, 500); }});
            document.addEventListener('DOMContentLoaded', function() {{ syncRightPanelTop(); }});
        }} else {{
            initCanvasControls();
            setTimeout(enhanceNumberInputs, 500);
            setTimeout(initPopoverDrag, 500);
            syncRightPanelTop();
        }}
        window.addEventListener('resize', syncRightPanelTop);

        // ── Esc key: close any open popover ──────────────────────
        document.addEventListener('keydown', function(e) {{
            if (e.key !== 'Escape') return;
            var popoverIds = [
                'file-details-popover', 'download-panel',
                'ordering-panel', 'guide-popover'
            ];
            popoverIds.forEach(function(id) {{
                var el = document.getElementById(id);
                if (el && el.style.display !== 'none') {{
                    el.style.display = 'none';
                }}
            }});
        }});
    }})();
    </script>
</body>
</html>
"""

    slice_options = [{"label": s, "value": i} for i, s in enumerate(meta0.get("slices", []))]

    level_state0 = {
        "show": {
            str(lv.get("key", "")): (str(lv.get("key", "")) == "group")
            for lv in (meta0.get("enclosure_levels", []) or [])
        },
        "last_category_col": meta0.get("category_col", None),
        "last_valid_keys": [str(lv.get("key", "")) for lv in (meta0.get("enclosure_levels", []) or []) if
                            str(lv.get("key", "")) != ""]
    }

    # Initialize supergroup style state so supergroup label positioning is correct BEFORE any slider interaction.
    supergroups_style0: Dict[str, Any] = {}
    for lv in (meta0.get("enclosure_levels", []) or []):
        k = str(lv.get("key", "") or "")
        if not k.startswith("supergroup"):
            continue
        lvl_str = k.replace("supergroup", "")
        try:
            lvl_int = int(lvl_str)
        except Exception:
            continue
        supergroups_style0[str(lvl_int)] = _get_default_supergroup_style(int(lvl_int))

    ctrl_label_style = {"fontSize": "40px", "whiteSpace": "nowrap", "flexShrink": "0"}
    ctrl_box_style = {"fontSize": f"{cfg.ui_font_size}px", "flexShrink": "1", "minWidth": "0"}

    # Right-panel two-column alignment (label column + control column)
    # Use the Group 'Radius' control row as the baseline width for the control column.
    right_label_w = "260px"
    right_ctrl_w = "520px"
    right_ctrl_col_style = {
        "display": "flex", "flexDirection": "row", "alignItems": "center",
        "justifyContent": "flex-start", "gap": "10px",
        "flex": "1", "minWidth": "0",
        "boxSizing": "border-box",
    }
    # For rows that contain ONLY one control (e.g., a single Dropdown), using a block container
    # avoids flex sizing quirks and ensures the control stretches to the full right edge.
    right_ctrl_single_style = {**right_ctrl_col_style, "display": "block"}

    # Some rows (e.g., show/hide name buttons) do not use the label+control columns.
    # Constrain them to the same total width as a standard "label + gap + control" row,
    # so their right edge aligns with other controls while keeping the left edge unchanged.
    right_toggle_row_w = f"calc({right_label_w} + {right_ctrl_w} + 10px)"
    # Two-input rows (Radius / Padding): 7-part grid alignment
    # Each line is split into 7 columns:
    #   label1 | gap1 | input1 | gap2 | label2 | gap3 | input2
    # Use fixed widths for label/input columns and let gap1/gap3 absorb the remaining space
    # so the second label and the right-most input align across all Radius/Padding rows.
    pair_lbl_w = "44px"
    # Interval tuning for two-input rows
    # gap1: between label1 and input1 (shrink)
    # gap2: between input1 and label2 (expand)
    # gap3: between label2 and input2 (shrink)
    pair_gap1_w = "8px"
    pair_gap2_w = "18px"
    pair_gap3_w = "8px"
    # Make both inputs stretch equally and stay perfectly aligned across all rows
    pair_input_col = "minmax(0, 1fr)"
    pair_two_input_row_style = {
        "display": "grid",
        "gridTemplateColumns": f"{pair_lbl_w} {pair_gap1_w} {pair_input_col} {pair_gap2_w} {pair_lbl_w} {pair_gap3_w} {pair_input_col}",
        "alignItems": "center",
        "width": "100%",
        "minWidth": "0",
        "boxSizing": "border-box",
    }
    pair_small_label_style = {**ctrl_label_style, "marginBottom": "0px", "width": "100%", "textAlign": "left"}
    pair_input_style = {"width": "100%", "minWidth": "0", "textAlign": "left", **ctrl_box_style}
    panel_style = {"padding": "10px", "border": "2px solid #AAAAAA", "borderRadius": "6px", "marginBottom": "10px",
                   "backgroundColor": "#FAFAFA", "overflow": "visible", "boxSizing": "border-box",
                   "position": "relative"}

    # --- Helper: build one On/Off + control row for block type ---
    def _blk_onoff_btn(suffix, bt):
        return html.Button(
            "Off",
            id={"type": f"blk-{suffix}-onoff", "index": bt},
            n_clicks=0,
            style={"minWidth": "48px", "fontSize": "28px", "padding": "2px 6px",
                   "borderRadius": "4px", "border": "2px solid #CCC",
                   "backgroundColor": "#F0F0F0", "cursor": "pointer"},
        )

    def _blk_color_row(label, suffix, bt, default_color, ctrl_label_style, ctrl_box_style):
        return html.Div(
            style={"display": "flex", "flexDirection": "row", "alignItems": "center", "gap": "8px",
                   "marginBottom": "8px"},
            children=[
                _blk_onoff_btn(suffix, bt),
                html.Label(label,
                           style={**ctrl_label_style, "marginBottom": "0px", "width": "auto", "fontSize": "34px"}),
                dcc.Input(debounce=True, id={"type": f"blk-{suffix}-picker-pt", "index": bt},
                          type="color", value=default_color,
                          style={"width": "40px", "height": "40px"}, disabled=True),
                dcc.Input(id={"type": f"blk-{suffix}-text-pt", "index": bt},
                          type="text", value=default_color, debounce=True,
                          style={"flex": "1", "minWidth": "0", **ctrl_box_style}, disabled=True),
            ],
        )

    def _blk_slider_row(label, suffix, bt, minv, maxv, step, default, ctrl_label_style, ctrl_box_style):
        return html.Div(
            style={"display": "flex", "flexDirection": "row", "alignItems": "center", "gap": "8px",
                   "marginBottom": "8px"},
            children=[
                _blk_onoff_btn(suffix, bt),
                html.Label(label,
                           style={**ctrl_label_style, "marginBottom": "0px", "width": "auto", "fontSize": "34px"}),
                html.Div(style={"flex": "1", "minWidth": "0"}, children=[
                    dcc.Slider(id={"type": f"blk-{suffix}-slider-pt", "index": bt},
                               min=minv, max=maxv, step=step, value=default,
                               updatemode="mouseup", tooltip={"always_visible": False}, disabled=True),
                ]),
                dcc.Input(debounce=True, id={"type": f"blk-{suffix}-input-pt", "index": bt},
                          type="number", min=minv, max=maxv, step=step, value=default,
                          style={"width": "80px", **ctrl_box_style}, disabled=True),
            ],
        )

    def _blk_dropdown_row(label, suffix, bt, options, default, ctrl_label_style, ctrl_box_style):
        return html.Div(
            style={"display": "flex", "flexDirection": "row", "alignItems": "center", "gap": "8px",
                   "marginBottom": "8px"},
            children=[
                _blk_onoff_btn(suffix, bt),
                html.Label(label,
                           style={**ctrl_label_style, "marginBottom": "0px", "width": "auto", "fontSize": "34px"}),
                html.Div(style={"flex": "1", "minWidth": "0"}, children=[
                    dcc.Dropdown(id={"type": f"blk-{suffix}-pt", "index": bt},
                                 options=[{"label": l, "value": v} for l, v in options],
                                 value=default, clearable=False,
                                 optionHeight=dropdown_option_h,
                                 style={**ctrl_box_style, "width": "100%"}, disabled=True),
                ]),
            ],
        )

    def _blk_radius_row(bt, ctrl_label_style, ctrl_box_style):
        pair_st = {"width": "60px", "fontSize": "28px", "textAlign": "center"}
        lbl_st = {"fontSize": "24px", "color": "#666", "flexShrink": "0"}
        return html.Div(
            style={"display": "flex", "flexDirection": "row", "alignItems": "center", "gap": "8px",
                   "marginBottom": "8px"},
            children=[
                _blk_onoff_btn("radius", bt),
                html.Label("Radius",
                           style={**ctrl_label_style, "marginBottom": "0px", "width": "auto", "fontSize": "34px"}),
                html.Div(style={"flex": "1", "minWidth": "0", "display": "flex", "flexDirection": "row",
                                "alignItems": "center", "justifyContent": "flex-end", "gap": "6px"}, children=[
                    html.Span("TL", style=lbl_st),
                    dcc.Input(id={"type": "blk-radius-tl-pt", "index": bt}, type="number",
                              min=0, max=50, step="any", value=0, debounce=False, style=pair_st, disabled=True),
                    html.Span("TR", style=lbl_st),
                    dcc.Input(id={"type": "blk-radius-tr-pt", "index": bt}, type="number",
                              min=0, max=50, step="any", value=0, debounce=False, style=pair_st, disabled=True),
                    html.Span("BL", style=lbl_st),
                    dcc.Input(id={"type": "blk-radius-bl-pt", "index": bt}, type="number",
                              min=0, max=50, step="any", value=0, debounce=False, style=pair_st, disabled=True),
                    html.Span("BR", style=lbl_st),
                    dcc.Input(id={"type": "blk-radius-br-pt", "index": bt}, type="number",
                              min=0, max=50, step="any", value=0, debounce=False, style=pair_st, disabled=True),
                ]),
            ],
        )

    # --- Helper: build one collapsible block-type section ---
    BIRTH_SUBTYPES = {"EnBirth", "ExBirth", "HyBirth"}

    def _build_block_type_section(bt, cfg, ctrl_label_style, ctrl_box_style):
        default_fill = cfg.block_colors.get(bt, "#CCCCCC")
        is_subtype = bt in BIRTH_SUBTYPES
        swatch_style = {
            "width": "16px", "height": "16px", "borderRadius": "3px",
            "backgroundColor": default_fill, "border": "2px solid #999",
            "marginRight": "8px", "flexShrink": "0",
        }
        header_padding = "6px 4px 6px 28px" if is_subtype else "6px 4px"
        return html.Div(
            style={"marginBottom": "2px"},
            children=[
                # Header
                html.Div(
                    style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                           "borderBottom": "2px solid #E0E0E0", "padding": header_padding},
                    children=[
                        html.Div(
                            id={"type": "blk-type-header", "index": bt},
                            n_clicks=0,
                            style={"display": "flex", "alignItems": "center", "cursor": "pointer",
                                   "flex": "1", "userSelect": "none"},
                            children=[
                                html.Span("▶", id={"type": "blk-type-arrow", "index": bt},
                                          style={"marginRight": "8px", "fontSize": "12px", "color": "#666",
                                                 "display": "inline-block", "width": "14px"}),
                                html.Div(style=swatch_style, id={"type": "blk-type-swatch", "index": bt}),
                                html.Span(bt, style={"fontSize": "36px", "fontWeight": "600"}),
                            ],
                        ),
                    ],
                ),
                # Collapsible body
                html.Div(
                    id={"type": "blk-type-body", "index": bt},
                    style={"visibility": "hidden", "height": "0", "overflow": "hidden", "padding": "8px 4px 8px 22px"},
                    children=[
                        _blk_color_row("Fill color", "fill-color", bt, default_fill, ctrl_label_style, ctrl_box_style),
                        _blk_slider_row("Fill opacity", "fill-opacity", bt, 0.0, 1.0, 0.01, 1.0, ctrl_label_style,
                                        ctrl_box_style),
                        _blk_slider_row("Border width", "border-width", bt, 0.0, 5.0, 0.1,
                                        float(cfg.block_border_width), ctrl_label_style, ctrl_box_style),
                        _blk_dropdown_row("Line style", "line-style", bt,
                                          [("Solid", "solid"), ("Dash", "dash"), ("Dot", "dot"),
                                           ("Dashdot", "dashdot")],
                                          "solid", ctrl_label_style, ctrl_box_style),
                        _blk_color_row("Border color", "border-color", bt, "#222222", ctrl_label_style, ctrl_box_style),
                        _blk_slider_row("Border opacity", "border-opacity", bt, 0.0, 1.0, 0.01, 1.0, ctrl_label_style,
                                        ctrl_box_style),
                        _blk_radius_row(bt, ctrl_label_style, ctrl_box_style),
                        _blk_dropdown_row("Text font", "text-font", bt,
                                          [("Arial", "Arial"), ("Helvetica", "Helvetica"),
                                           ("Times New Roman", "Times New Roman"),
                                           ("Courier New", "Courier New"), ("Georgia", "Georgia"),
                                           ("Verdana", "Verdana"),
                                           ("Microsoft YaHei", "Microsoft YaHei"), ("SimHei", "SimHei"),
                                           ("SimSun", "SimSun")],
                                          "Arial", ctrl_label_style, ctrl_box_style),
                        _blk_slider_row("Text size", "text-size", bt, 8, 72, 1, int(cfg.block_text_size),
                                        ctrl_label_style, ctrl_box_style),
                        _blk_color_row("Text color", "text-color", bt, cfg.block_text_colors.get(bt, "#111111"),
                                       ctrl_label_style, ctrl_box_style),
                        _blk_dropdown_row("Text align", "text-align", bt,
                                          [("Left", "left"), ("Center", "center"), ("Right", "right")],
                                          "center", ctrl_label_style, ctrl_box_style),
                        _blk_slider_row("Line spacing", "line-spacing", bt, -20, 200, 1, 0, ctrl_label_style,
                                        ctrl_box_style),
                        _blk_slider_row("Text rotation", "text-rotation", bt, -180, 180, 1, 90, ctrl_label_style,
                                        ctrl_box_style),
                    ],
                ),
            ],
        )

    def _build_block_all_section(cfg, ctrl_label_style, ctrl_box_style):
        """Build a special 'All' section with only text-related controls (Text font through Text rotation).
        Hidden dummy inputs are included for the 7 non-text controls so that pattern-matching
        alignment with other block types remains consistent."""
        bt = "All"
        swatch_style = {
            "width": "16px", "height": "16px", "borderRadius": "3px",
            "backgroundColor": "#888888", "border": "2px solid #999",
            "marginRight": "8px", "flexShrink": "0",
        }
        # Hidden dummy inputs for non-text controls (keeps pattern-matching ALL lists aligned)
        hidden_dummies = html.Div(style={"display": "none"}, children=[
            dcc.Input(id={"type": "blk-fill-color-picker-pt", "index": bt}, type="color", value="#CCCCCC"),
            dcc.Input(id={"type": "blk-fill-color-text-pt", "index": bt}, type="text", value="#CCCCCC"),
            html.Button("Off", id={"type": "blk-fill-color-onoff", "index": bt}, n_clicks=0),
            dcc.Slider(id={"type": "blk-fill-opacity-slider-pt", "index": bt}, min=0, max=1, step=0.01, value=1.0),
            dcc.Input(id={"type": "blk-fill-opacity-input-pt", "index": bt}, type="number", value=1.0),
            html.Button("Off", id={"type": "blk-fill-opacity-onoff", "index": bt}, n_clicks=0),
            dcc.Slider(id={"type": "blk-border-width-slider-pt", "index": bt}, min=0, max=3, step=0.05,
                       value=float(cfg.block_border_width)),
            dcc.Input(id={"type": "blk-border-width-input-pt", "index": bt}, type="number",
                      value=float(cfg.block_border_width)),
            html.Button("Off", id={"type": "blk-border-width-onoff", "index": bt}, n_clicks=0),
            dcc.Dropdown(id={"type": "blk-line-style-pt", "index": bt}, options=[{"label": "Solid", "value": "solid"}],
                         value="solid"),
            html.Button("Off", id={"type": "blk-line-style-onoff", "index": bt}, n_clicks=0),
            dcc.Input(id={"type": "blk-border-color-picker-pt", "index": bt}, type="color", value="#222222"),
            dcc.Input(id={"type": "blk-border-color-text-pt", "index": bt}, type="text", value="#222222"),
            html.Button("Off", id={"type": "blk-border-color-onoff", "index": bt}, n_clicks=0),
            dcc.Slider(id={"type": "blk-border-opacity-slider-pt", "index": bt}, min=0, max=1, step=0.01, value=1.0),
            dcc.Input(id={"type": "blk-border-opacity-input-pt", "index": bt}, type="number", value=1.0),
            html.Button("Off", id={"type": "blk-border-opacity-onoff", "index": bt}, n_clicks=0),
            dcc.Input(id={"type": "blk-radius-tl-pt", "index": bt}, type="number", min=0, max=50, step="any", value=0),
            dcc.Input(id={"type": "blk-radius-tr-pt", "index": bt}, type="number", min=0, max=50, step="any", value=0),
            dcc.Input(id={"type": "blk-radius-bl-pt", "index": bt}, type="number", min=0, max=50, step="any", value=0),
            dcc.Input(id={"type": "blk-radius-br-pt", "index": bt}, type="number", min=0, max=50, step="any", value=0),
            html.Button("Off", id={"type": "blk-radius-onoff", "index": bt}, n_clicks=0),
        ])
        return html.Div(
            style={"marginBottom": "2px"},
            children=[
                # Header
                html.Div(
                    style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                           "borderBottom": "2px solid #999", "padding": "6px 4px",
                           "backgroundColor": "#F0F0F0"},
                    children=[
                        html.Div(
                            id={"type": "blk-type-header", "index": bt},
                            n_clicks=0,
                            style={"display": "flex", "alignItems": "center", "cursor": "pointer",
                                   "flex": "1", "userSelect": "none"},
                            children=[
                                html.Span("▶", id={"type": "blk-type-arrow", "index": bt},
                                          style={"marginRight": "8px", "fontSize": "12px", "color": "#666",
                                                 "display": "inline-block", "width": "14px"}),
                                html.Div(style=swatch_style, id={"type": "blk-type-swatch", "index": bt}),
                                html.Span("All", style={"fontSize": "36px", "fontWeight": "700"}),
                            ],
                        ),
                    ],
                ),
                # Collapsible body — only text controls
                html.Div(
                    id={"type": "blk-type-body", "index": bt},
                    style={"visibility": "hidden", "height": "0", "overflow": "hidden", "padding": "8px 4px 8px 22px"},
                    children=[
                        _blk_dropdown_row("Text font", "text-font", bt,
                                          [("Arial", "Arial"), ("Helvetica", "Helvetica"),
                                           ("Times New Roman", "Times New Roman"),
                                           ("Courier New", "Courier New"), ("Georgia", "Georgia"),
                                           ("Verdana", "Verdana"),
                                           ("Microsoft YaHei", "Microsoft YaHei"), ("SimHei", "SimHei"),
                                           ("SimSun", "SimSun")],
                                          "Arial", ctrl_label_style, ctrl_box_style),
                        _blk_slider_row("Text size", "text-size", bt, 8, 72, 1, int(cfg.block_text_size),
                                        ctrl_label_style, ctrl_box_style),
                        _blk_color_row("Text color", "text-color", bt, "#111111", ctrl_label_style, ctrl_box_style),
                        _blk_dropdown_row("Text align", "text-align", bt,
                                          [("Left", "left"), ("Center", "center"), ("Right", "right")],
                                          "center", ctrl_label_style, ctrl_box_style),
                        _blk_slider_row("Line spacing", "line-spacing", bt, -20, 200, 1, 0, ctrl_label_style,
                                        ctrl_box_style),
                        _blk_slider_row("Text rotation", "text-rotation", bt, -180, 180, 1, 90, ctrl_label_style,
                                        ctrl_box_style),
                    ],
                ),
                # Hidden dummies for non-text controls
                hidden_dummies,
            ],
        )

    # --- Helper: build one collapsible band-type section ---
    def _build_band_type_section(bt: str, cfg, ctrl_label_style, ctrl_box_style, right_ctrl_col_style, right_label_w):
        """Build a collapsible section for one band type with On/Off, Color, Opacity controls."""
        default_color = cfg.band_colors.get(bt, "#999999")
        default_opacity = float(cfg.band_default_opacity)

        # Color swatch shown in header
        swatch_style = {
            "width": "16px", "height": "16px", "borderRadius": "3px",
            "backgroundColor": default_color, "border": "2px solid #999",
            "marginRight": "8px", "flexShrink": "0",
        }

        hide_btn_style = {
            "fontSize": "24px", "padding": "2px 10px", "borderRadius": "4px",
            "border": "2px solid #CCC", "backgroundColor": "#FFF", "color": "#333",
            "cursor": "pointer", "width": "100px", "flexShrink": "0",
            "lineHeight": "1.4", "textAlign": "center",
        }

        return html.Div(
            style={"marginBottom": "2px"},
            children=[
                # Header row: left = collapsible click area, right = hide/show button
                html.Div(
                    style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                           "borderBottom": "2px solid #E0E0E0", "padding": "6px 4px"},
                    children=[
                        # Left: clickable area for collapse
                        html.Div(
                            id={"type": "band-type-header", "index": bt},
                            n_clicks=0,
                            style={"display": "flex", "alignItems": "center", "cursor": "pointer",
                                   "flex": "1", "userSelect": "none"},
                            children=[
                                html.Span("▶", id={"type": "band-type-arrow", "index": bt},
                                          style={"marginRight": "8px", "fontSize": "12px", "color": "#666",
                                                 "transition": "transform 0.15s", "display": "inline-block",
                                                 "width": "14px"}),
                                html.Div(style=swatch_style, id={"type": "band-type-swatch", "index": bt}),
                                html.Span(bt, style={"fontSize": "36px", "fontWeight": "600"}),
                            ],
                        ),
                        # Right: per-type Hide/Show button
                        html.Button(
                            "Hide",
                            id={"type": "band-type-hide-btn", "index": bt},
                            n_clicks=0,
                            style=hide_btn_style,
                        ),
                    ],
                ),
                # Collapsible body (hidden by default)
                html.Div(
                    id={"type": "band-type-body", "index": bt},
                    style={"visibility": "hidden", "height": "0", "overflow": "hidden", "padding": "8px 4px 8px 22px"},
                    children=[
                        # --- Color row ---
                        html.Div(
                            style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                                   "gap": "8px", "marginBottom": "8px"},
                            children=[
                                html.Button(
                                    id={"type": "band-color-onoff", "index": bt},
                                    n_clicks=0,
                                    style={"minWidth": "48px", "fontSize": "28px", "padding": "2px 6px",
                                           "borderRadius": "4px", "border": "2px solid #CCC",
                                           "backgroundColor": "#F0F0F0", "cursor": "pointer"},
                                ),
                                html.Label("Color", style={**ctrl_label_style, "marginBottom": "0px", "width": "auto",
                                                           "fontSize": "34px"}),
                                dcc.Input(debounce=True,
                                          id={"type": "band-color-picker-pt", "index": bt},
                                          type="color",
                                          value=default_color,
                                          style={"width": "40px", "height": "40px"},
                                          disabled=True,
                                          ),
                                dcc.Input(
                                    id={"type": "band-color-text-pt", "index": bt},
                                    type="text",
                                    value=default_color,
                                    debounce=True,
                                    style={"flex": "1", "minWidth": "0", **ctrl_box_style},
                                    disabled=True,
                                ),
                            ],
                        ),
                        # --- Opacity row ---
                        html.Div(
                            style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                                   "gap": "8px"},
                            children=[
                                html.Button(
                                    "Off",
                                    id={"type": "band-opacity-onoff", "index": bt},
                                    n_clicks=0,
                                    style={"minWidth": "48px", "fontSize": "28px", "padding": "2px 6px",
                                           "borderRadius": "4px", "border": "2px solid #CCC",
                                           "backgroundColor": "#F0F0F0", "cursor": "pointer"},
                                ),
                                html.Label("Opacity", style={**ctrl_label_style, "marginBottom": "0px", "width": "auto",
                                                             "fontSize": "34px"}),
                                html.Div(
                                    style={"flex": "1", "minWidth": "0"},
                                    children=[
                                        dcc.Slider(
                                            id={"type": "band-opacity-slider-pt", "index": bt},
                                            min=0.0, max=1.0, step=0.01,
                                            value=default_opacity,
                                            updatemode="mouseup",
                                            tooltip={"always_visible": False},
                                            disabled=True,
                                        ),
                                    ],
                                ),
                                dcc.Input(debounce=True,
                                          id={"type": "band-opacity-input-pt", "index": bt},
                                          type="number",
                                          min=0.0, max=1.0, step="any",
                                          value=default_opacity,
                                          style={"width": "80px", **ctrl_box_style},
                                          disabled=True,
                                          ),
                            ],
                        ),
                    ],
                ),
            ],
        )

    app.layout = html.Div([
        dcc.Location(id="url", refresh=False),
        # ========== TOP BAR ==========
        # ========== GRAPH (full-window background layer) ==========
        # Fills the entire viewport. Top-bar and right-panel overlay on top via z-index.
        # Initial width uses the SAME formula as applyUITransform so there is no
        # first-frame jump: panel visual width = 920 * uiScale, where
        # uiScale = clamp(innerWidth / (6*920), 0.15, 2.0). We approximate this with
        # pure CSS: visual_panel = clamp(138px, 100vw/6, 1840px).
        html.Div(
            id="graph-wrap",
            style={
                "position": "fixed",
                "top": "84px",
                "left": "0",
                "width": "calc(100vw - clamp(138px, 100vw / 6, 1840px))",
                "height": "calc(100vh - 84px)",
                "zIndex": "0",
                "overflow": "hidden",   # clips CSS-transformed canvas content
            },
            children=[
                html.Button(id="btn-wrap-leftclick", n_clicks=0, style={"display": "none"}),
                # canvas-content: the CSS-transform target for zoom/pan.
                # Everything inside scales as one unit — no Plotly recalculation.
                html.Div(
                    id="canvas-content",
                    style={
                        "position": "absolute",
                        "top": "0", "left": "0",
                        "width": "100%", "height": "100%",
                        "transformOrigin": "0 0",
                        "willChange": "transform",
                    },
                    children=[
                        dcc.Graph(
                            id="graph",
                            figure=fig0,
                            style={"width": "100%", "height": "100%"},
                            responsive=True,
                            config={
                                "responsive": True,
                                "displaylogo": False,
                                "doubleClick": False,
                                "scrollZoom": False,
                                "dragmode": False,
                                "modeBarButtonsToRemove": [
                                    "zoom2d", "zoomIn2d", "zoomOut2d", "autoScale2d",
                                    "select2d", "lasso2d", "pan2d",
                                    "toImage", "resetScale2d", "hoverClosestCartesian", "hoverCompareCartesian"
                                ],
                            },
                        ),
                        # HTML-based no-data hint: lives outside Plotly's SVG so
                        # it cannot be reset by Plotly re-renders. JS controls
                        # its size (counter-scaled against canvasZoom) and
                        # show/hide (based on _meta.init_x_span).
                        html.Div(
                            [
                                html.Span("No data loaded.", style={"display": "block"}),
                                html.Span("Click 'Upload files'", style={"display": "block"}),
                                html.Span("to select one or more slice tables.",
                                          style={"display": "block"}),
                            ],
                            id="no-data-hint-html",
                            style={
                                "display": "block",
                                "position": "absolute",
                                "top": "50%", "left": "50%",
                                # Include scale(1) in the initial transform so it has
                                # exactly the same structure as the JS-written value:
                                #   translate(-50%,-50%) scale(1/canvasZoom)
                                # This avoids a repaint/resize flicker the first time
                                # applyCanvasTransform() runs (observed in Safari/Firefox).
                                "transform": "translate(-50%, -50%) scale(1)",
                                "transformOrigin": "center center",
                                "fontSize": "0.75vw",
                                "textAlign": "center",
                                "color": "#666666",
                                "lineHeight": "1.7",
                                "pointerEvents": "none",
                                "zIndex": "10",
                                "whiteSpace": "nowrap",
                            },
                        ),
                    ],
                ),
                html.Div(
                    id="graph-loading-overlay",
                    className="loading-overlay",
                    children=[html.Div(className="loading-spinner")],
                ),
                html.Div(
                    "Ctrl+Scroll: Zoom | Middle-drag: Pan",
                    id="canvas-zoom-hint",
                    style={
                        "position": "absolute",
                        "bottom": "8px",
                        "left": "8px",
                        # 0.75vw matches the JS formula in applyUITransform
                        # (_hintPx = innerWidth * 0.0075). Use the same unit at
                        # first render so there is no visible size jump when JS
                        # takes over. Do NOT hard-code "20px" — that caused a
                        # jarring shrink on load at common viewport widths.
                        "fontSize": "0.75vw",
                        "color": "#888888",
                        "backgroundColor": "rgba(255,255,255,0.8)",
                        "padding": "2px 6px",
                        "borderRadius": "3px",
                        "pointerEvents": "none",
                        "zIndex": "1000",
                    }
                ),
            ],
        ),

        html.Div(id="ui-top-wrap", children=[
            html.Div(id="top-bar", children=[
                html.Div(
                    style={"display": "flex", "alignItems": "center", "height": "64px", "position": "relative"},
                    children=[
                        dcc.Upload(
                            id="upload-slices",
                            children=html.Button(
                                "Upload files",
                                title="Upload one multi-sheet Excel file, or select multiple files (order dialog will appear).",
                            ),
                            multiple=True,
                            accept=".xlsx,.xls,.csv,.ods",
                            style={"display": "flex"},
                        ),
                        dcc.Store(id="store-pending-files", data={"paths": [], "names": []}),
                        html.Div(
                            style={"position": "relative", "display": "inline-block", "marginLeft": "10px"},
                            children=[
                                html.Button("File details", id="btn-file-details", n_clicks=0),

                                html.Div(
                                    id="file-details-popover",
                                    style={
                                        "display": "none",
                                        "position": "absolute",
                                        "top": "100%",
                                        "left": "0px",
                                        "marginTop": "6px",
                                        "zIndex": "30000",
                                        "backgroundColor": "#FFFFFF",
                                        "border": "2px solid #CCCCCC",
                                        "borderRadius": "8px",
                                        "padding": "30px 30px",
                                        "minWidth": "440px",
                                        "boxShadow": "0 6px 18px rgba(0,0,0,0.18)",
                                        "fontSize": "40px",
                                        "whiteSpace": "pre-line",
                                        "cursor": "move",
                                    },

                                    children=[
                                        html.Div(
                                            style={"display": "flex", "justifyContent": "flex-end",
                                                   "alignItems": "center"},
                                            children=[
                                                html.Span(
                                                    "×",
                                                    id="btn-file-details-close",
                                                    n_clicks=0,
                                                    style={
                                                        "fontSize": "40px",
                                                        "lineHeight": "40px",
                                                        "cursor": "pointer",
                                                        "userSelect": "none",
                                                        "padding": "0px",
                                                        "margin": "0px",
                                                        "border": "none",
                                                        "background": "transparent",
                                                    },
                                                ),
                                            ],
                                        ),
                                        html.Div(id="file-details-popover-text",
                                                 style={"marginTop": "6px", "width": "100%", "textAlign": "left"}),
                                    ],
                                ),
                            ],
                        ),

                        # --- Initial combo width ---
                        # Uses the same formula as the
                        # `update_category_options_from_upload` callback, so
                        # the dropdown can display the longest option name
                        # from the moment the page loads (not just after
                        # upload). Without this, on initial load long level
                        # names like "Hierarchy 4" / "Supergroup" could be
                        # clipped when the dropdown menu opened.
                        html.Div(
                            id="category-combo",
                            className="category-combo-btn",
                            style={
                                "marginLeft": "18px",
                                # No minWidth here — the outer combo's width
                                # is determined by its children (label +
                                # wrapper div), and the wrapper div's width
                                # is locked by `update_category_options_from_upload`
                                # to fit the longest option.
                                "flex": "0 0 auto",
                            },
                            children=[
                                html.Div("Category selection", className="category-combo-label"),
                                # Wrap dcc.Dropdown in a fixed-width div so the
                                # outer combo's width never depends on the
                                # internal react-select layout.
                                #
                                # Why: dcc.Dropdown wraps react-select v1, which
                                # is well-known to not have a built-in option to
                                # lock its width — it sizes its inner DOM around
                                # the currently selected value. As that value
                                # changes (user picks a different level), the
                                # internal width changes too, and the outer
                                # inline-flex .category-combo-btn shrinks/grows
                                # to fit.
                                # See: https://github.com/JedWatson/react-select/issues/4201
                                #
                                # Solution: this wrapper Div has a fixed width
                                # set via inline style. Inside it the dropdown
                                # is set to width:100%, so it expands to fill
                                # the wrapper. react-select can do whatever it
                                # wants internally — the wrapper's width stays
                                # constant, so the whole .category-combo-btn
                                # also stays constant.
                                html.Div(
                                    id="category-col-wrap",
                                    style={
                                        "flex": "1 1 auto",
                                        "minWidth": "300px",
                                        "alignSelf": "stretch",
                                    },
                                    children=[
                                        dcc.Dropdown(
                                            id="category-col",
                                            options=[{"label": c, "value": c} for c in category_candidates],
                                            value=category_col,
                                            clearable=False,
                                            optionHeight=dropdown_option_h,
                                            className="category-combo-dropdown",
                                            style={"width": "100%"},
                                        ),
                                    ],
                                ),
                            ],
                        ),

                        html.Div(
                            style={"position": "relative", "display": "inline-block", "marginLeft": "10px"},
                            children=[
                                html.Button(
                                    "Download",
                                    id="btn-download",
                                    n_clicks=0,
                                ),
                                dcc.Download(id="download-image"),
                                dcc.Store(id="download-config"),
                                html.Div(id="dl-dummy", style={"display": "none"}),
                                dcc.Store(id="__resize_trigger"),
                                html.Div(id="print-dummy", style={"display": "none"}),

                                html.Div(
                                    id="download-panel",
                                    style={
                                        "display": "none",
                                        "position": "absolute",
                                        "top": "100%",
                                        "left": "0px",
                                        "marginTop": "6px",
                                        "zIndex": "30000",
                                        "backgroundColor": "#FFFFFF",
                                        "border": "2px solid #CCCCCC",
                                        "borderRadius": "8px",
                                        "padding": "30px 30px",
                                        "width": "700px",
                                        "boxSizing": "border-box",
                                        "boxShadow": "0 6px 18px rgba(0,0,0,0.18)",
                                        "fontSize": "40px",
                                        "cursor": "move",
                                    },
                                    children=[
                                        html.Div(
                                            id="download-panel-inner",
                                            style={"width": "640px"},
                                            children=[
                                                html.Div(
                                                    style={"display": "flex", "justifyContent": "flex-end",
                                                           "alignItems": "center"},
                                                    children=[
                                                        html.Span(
                                                            "×",
                                                            id="btn-download-close",
                                                            n_clicks=0,
                                                            style={
                                                                "fontSize": "40px",
                                                                "lineHeight": "40px",
                                                                "cursor": "pointer",
                                                                "userSelect": "none",
                                                                "padding": "0px",
                                                                "margin": "0px",
                                                                "border": "none",
                                                                "background": "transparent",
                                                            },
                                                        ),
                                                    ],
                                                ),

                                                html.Div(
                                                    style={"display": "flex", "flexDirection": "row",
                                                           "alignItems": "center", "justifyContent": "space-between",
                                                           "marginTop": "20px", "width": "640px"},
                                                    children=[
                                                        html.Label("File name",
                                                                   style={**ctrl_label_style, "marginBottom": "0px",
                                                                          "textAlign": "left"}),
                                                        dcc.Input(debounce=True, id="dl-name", value="birdcage_diagram",
                                                                  type="text",
                                                                  style={"width": "300px", "boxSizing": "border-box",
                                                                         **ctrl_box_style}),
                                                    ],
                                                ),

                                                html.Div(
                                                    style={"display": "flex", "flexDirection": "row",
                                                           "alignItems": "center", "justifyContent": "space-between",
                                                           "marginTop": "20px", "width": "640px"},
                                                    children=[
                                                        html.Label("Format",
                                                                   style={**ctrl_label_style, "marginBottom": "0px",
                                                                          "textAlign": "left"}),
                                                        dcc.Dropdown(
                                                            id="dl-format",
                                                            options=[
                                                                {"label": "PNG", "value": "png"},
                                                                {"label": "JPEG", "value": "jpeg"},
                                                                {"label": "SVG", "value": "svg"},
                                                            ],
                                                            value="png",
                                                            clearable=False,
                                                            optionHeight=dropdown_option_h,
                                                            style={"width": "300px", "minWidth": "300px",
                                                                   "maxWidth": "300px", "boxSizing": "border-box"},
                                                        ),
                                                    ],
                                                ),

                                                html.Div(
                                                    style={"display": "flex", "flexDirection": "row",
                                                           "alignItems": "center", "justifyContent": "space-between",
                                                           "marginTop": "20px", "width": "640px"},
                                                    children=[
                                                        html.Label("Resolution (DPI)",
                                                                   style={**ctrl_label_style, "marginBottom": "0px",
                                                                          "textAlign": "left"}),
                                                        dcc.Input(debounce=True, id="dl-dpi", type="number", value=300,
                                                                  min=1, step="any",
                                                                  style={"width": "300px", "boxSizing": "border-box",
                                                                         **ctrl_box_style}),
                                                    ],
                                                ),

                                                html.Div(
                                                    style={"display": "flex", "flexDirection": "row",
                                                           "alignItems": "center", "justifyContent": "space-between",
                                                           "marginTop": "20px", "width": "640px"},
                                                    children=[
                                                        html.Label("Canvas scale",
                                                                   style={**ctrl_label_style, "marginBottom": "0px",
                                                                          "textAlign": "left"}),
                                                        dcc.Dropdown(
                                                            id="dl-scale",
                                                            options=[
                                                                {"label": "1×", "value": 1},
                                                                {"label": "2×", "value": 2},
                                                                {"label": "3×", "value": 3},
                                                            ],
                                                            value=1,
                                                            clearable=False,
                                                            optionHeight=dropdown_option_h,
                                                            style={"width": "300px", "minWidth": "300px",
                                                                   "maxWidth": "300px", "boxSizing": "border-box"},
                                                        ),
                                                    ],
                                                ),

                                                html.Button("Confirm download", id="btn-download-confirm", n_clicks=0,
                                                            style={"marginTop": "20px", "width": "640px"}),
                                            ],
                                        ),
                                    ],
                                ),
                            ],
                        ),

                        # Ordering parameter button and popover
                        html.Div(
                            style={"position": "relative", "display": "inline-block", "marginLeft": "10px"},
                            children=[
                                html.Button(
                                    "Ordering",
                                    id="btn-ordering",
                                    n_clicks=0,
                                ),

                                html.Div(
                                    id="ordering-panel",
                                    style={
                                        "display": "none",
                                        "position": "absolute",
                                        "top": "100%",
                                        "left": "0px",
                                        "marginTop": "6px",
                                        "zIndex": "30000",
                                        "backgroundColor": "#FFFFFF",
                                        "border": "2px solid #CCCCCC",
                                        "borderRadius": "8px",
                                        "padding": "30px 30px",
                                        "width": "900px",
                                        "boxSizing": "border-box",
                                        "boxShadow": "0 6px 18px rgba(0,0,0,0.18)",
                                        "fontSize": "40px",
                                        "cursor": "move",
                                    },
                                    children=[
                                        html.Div(
                                            id="ordering-panel-inner",
                                            style={"width": "840px"},
                                            children=[
                                                html.Div(
                                                    style={"display": "flex", "justifyContent": "flex-end",
                                                           "alignItems": "center"},
                                                    children=[
                                                        html.Span(
                                                            "×",
                                                            id="btn-ordering-close",
                                                            n_clicks=0,
                                                            style={
                                                                "fontSize": "40px",
                                                                "lineHeight": "40px",
                                                                "cursor": "pointer",
                                                                "userSelect": "none",
                                                                "padding": "0px",
                                                                "margin": "0px",
                                                                "border": "none",
                                                                "background": "transparent",
                                                            },
                                                        ),
                                                    ],
                                                ),

                                                html.Div(
                                                    style={"display": "flex", "flexDirection": "row",
                                                           "alignItems": "center", "justifyContent": "space-between",
                                                           "marginTop": "20px", "width": "840px"},
                                                    children=[
                                                        html.Label("Max iterations",
                                                                   style={**ctrl_label_style, "marginBottom": "0px",
                                                                          "textAlign": "left", "flex": "1",
                                                                          "minWidth": "0"}),
                                                        dcc.Input(debounce=True, id="sweep-k-max", type="number",
                                                                  value=10, min=1, max=100, step="any",
                                                                  style={"width": "300px", "minWidth": "300px",
                                                                         "boxSizing": "border-box", **ctrl_box_style}),
                                                    ],
                                                ),

                                                html.Div(
                                                    style={"display": "flex", "flexDirection": "row",
                                                           "alignItems": "center", "justifyContent": "space-between",
                                                           "marginTop": "20px", "width": "840px"},
                                                    children=[
                                                        html.Label("Consecutive stable rounds",
                                                                   style={**ctrl_label_style, "marginBottom": "0px",
                                                                          "textAlign": "left", "flex": "1",
                                                                          "minWidth": "0"}),
                                                        dcc.Input(debounce=True, id="sweep-m", type="number", value=2,
                                                                  min=1, max=20, step="any",
                                                                  style={"width": "300px", "minWidth": "300px",
                                                                         "boxSizing": "border-box", **ctrl_box_style}),
                                                    ],
                                                ),

                                                html.Div(
                                                    style={"display": "flex", "flexDirection": "row",
                                                           "alignItems": "center", "justifyContent": "space-between",
                                                           "marginTop": "20px", "width": "840px"},
                                                    children=[
                                                        html.Label("Convergence threshold",
                                                                   style={**ctrl_label_style, "marginBottom": "0px",
                                                                          "textAlign": "left", "flex": "1",
                                                                          "minWidth": "0"}),
                                                        dcc.Input(debounce=True, id="sweep-delta", type="number",
                                                                  value=0.01, min=0.001, max=1, step="any",
                                                                  style={"width": "300px", "minWidth": "300px",
                                                                         "boxSizing": "border-box", **ctrl_box_style}),
                                                    ],
                                                ),

                                                html.Button("Confirm", id="btn-ordering-confirm", n_clicks=0,
                                                            style={"marginTop": "20px", "width": "840px"}),
                                            ],
                                        ),
                                    ],
                                ),
                            ],
                        ),

                        # Guide button and popover
                        html.Div(
                            style={"position": "relative", "display": "inline-block", "marginLeft": "10px"},
                            children=[
                                html.Button("Guide", id="btn-guide", n_clicks=0),

                                html.Div(
                                    id="guide-popover",
                                    style={
                                        "display": "none",
                                        "position": "absolute",
                                        "top": "100%",
                                        "left": "0px",
                                        "marginTop": "6px",
                                        "zIndex": "30000",
                                        "backgroundColor": "#FFFFFF",
                                        "border": "2px solid #CCCCCC",
                                        "borderRadius": "8px",
                                        "padding": "30px 30px",
                                        "minWidth": "440px",
                                        "boxShadow": "0 6px 18px rgba(0,0,0,0.18)",
                                        "fontSize": "40px",
                                        "whiteSpace": "pre-line",
                                        "cursor": "move",
                                    },

                                    children=[
                                        html.Div(
                                            style={"display": "flex", "justifyContent": "flex-end",
                                                   "alignItems": "center"},
                                            children=[
                                                html.Span(
                                                    "×",
                                                    id="btn-guide-close",
                                                    n_clicks=0,
                                                    style={
                                                        "fontSize": "40px",
                                                        "lineHeight": "40px",
                                                        "cursor": "pointer",
                                                        "userSelect": "none",
                                                        "padding": "0px",
                                                        "margin": "0px",
                                                        "border": "none",
                                                        "background": "transparent",
                                                    },
                                                ),
                                            ],
                                        ),
                                        html.Div(id="guide-popover-text", children=[
                                            html.P([
                                                html.B("User Guide: "),
                                                html.A(
                                                    "Birdcage_Diagram_Use_Guide.pdf",
                                                    href="https://github.com/zhongjiang-licely/Birdcage-Diagram/raw/main/assets/Birdcage_Diagram_Use_Guide.pdf",
                                                    target="_blank",
                                                    style={"color": "#1a73e8", "textDecoration": "underline"},
                                                ),
                                            ]),
                                        ],
                                                 style={"marginTop": "6px", "width": "100%", "textAlign": "left"}),
                                    ],
                                ),
                            ],
                        ),

                        # Updating status badge — sits inline with toolbar
                        # buttons (next to Guide), inherits the toolbar's
                        # font size. Visibility / label text are controlled
                        # by the loading clientside callback.
                        html.Div(
                            id="updating-status",
                            children=[
                                html.Span("●", className="updating-dot"),
                                html.Span("Updating…", className="updating-label"),
                            ],
                        ),
                    ],
                ),

                html.Div(
                    id="upload-status",
                    style={"fontSize": f"{cfg.ui_font_size - 4}px", "color": "#555555", "marginTop": "4px"}),

                html.Div(),
            ], style={"padding": "12px 17px 7px 17px", "flexShrink": "0"}),

        ], style={
            "position": "fixed",
            "top": "0",
            "left": "0",
            "right": "0",
            "zIndex": "20000",
            "backgroundColor": "white",
            "borderBottom": "2px solid #DDDDDD",
        }),

        html.Div(
            id="right-panel",
            style={
                "position": "fixed",
                "top": "0",              # Always 0 — internal spacer aligns content below top-bar
                "right": "0",
                "width": "920px",
                "minWidth": "920px",
                "maxWidth": "920px",
                "height": "100%",        # applyUITransform overrides to 100/uiScale % for correct visual height
                "paddingTop": "0",
                "paddingBottom": "24px",
                "paddingLeft": "20px",
                "paddingRight": "20px",
                "borderLeft": "2px solid #DDDDDD",
                "overflowY": "auto",
                "overflowX": "hidden",
                "display": "flex",
                "flexDirection": "column",
                "gap": "0px",
                "boxSizing": "border-box",
                "backgroundColor": "white",
                "zIndex": "100",
            },

            children=[
                # Spacer whose layout height equals the top-bar's layout height (84px).
                # Both are scaled by the same uiScale, so their visual heights always match —
                # the right-panel content therefore always starts right below the top-bar,
                # regardless of uiScale. JS updates this height if the top-bar grows
                # (e.g. upload-status appears).
                html.Div(id="right-panel-topspacer",
                         style={"height": "84px", "minHeight": "84px",
                                "flexShrink": "0", "pointerEvents": "none"}),
                html.Div(
                    id="right-panel-inner",
                    style={
                        "width": "100%",
                        "boxSizing": "border-box",
                        "display": "flex",
                        "flexDirection": "column",
                        "alignItems": "stretch",
                    },
                    children=[
                        # ========== LAYOUT + LABEL SETTINGS (MERGED TOP PANEL) ==========
                        html.Div([
                            html.Div(
                                style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                                       "justifyContent": "flex-start", "gap": "10px", "width": "100%"},
                                children=[
                                    html.Label("Layer gap",
                                               style={**ctrl_label_style, "marginBottom": "0px", "width": "260px"}),
                                    html.Div(
                                        style=right_ctrl_col_style,
                                        children=[
                                            html.Div(
                                                style={"flex": "1", "minWidth": "0"},
                                                children=[
                                                    dcc.Slider(
                                                        id="layer-gap-slider",
                                                        min=1.0,
                                                        max=20.0,
                                                        step=0.1,
                                                        value=float(cfg.layer_gap),
                                                        updatemode="mouseup", tooltip={"always_visible": False},
                                                    ),
                                                ],
                                            ),
                                            dcc.Input(debounce=True,
                                                      id="layer-gap-input",
                                                      type="number",
                                                      min=1.0,
                                                      max=20.0,
                                                      step="any",
                                                      value=float(cfg.layer_gap),
                                                      style={"width": "100px", **ctrl_box_style},
                                                      ),
                                        ],
                                    ),
                                ],
                            ),

                            html.Div(
                                style={"marginTop": "10px", "boxSizing": "border-box"},
                                children=[
                                    html.Button(
                                        "Show Exo",
                                        id="btn-exo-all",
                                        n_clicks=0,
                                        style={"width": "100%", "whiteSpace": "nowrap"},
                                    ),
                                ],
                            ),

                            html.Div(
                                style={"display": "flex", "flexDirection": "row", "alignItems": "center", "gap": "10px",
                                       "marginTop": "10px", "width": "100%"},
                                children=[
                                    html.Label("Exo distance",
                                               style={**ctrl_label_style, "marginBottom": "0px", "width": "260px"}),
                                    html.Div(
                                        style=right_ctrl_col_style,
                                        children=[
                                            html.Div(
                                                style={"flex": "1", "minWidth": "0"},
                                                children=[
                                                    dcc.Slider(
                                                        id="exo-distance-slider",
                                                        min=0.5,
                                                        max=5.0,
                                                        step=0.1,
                                                        value=float(cfg.exo_gap),
                                                        updatemode="mouseup", tooltip={"always_visible": False},
                                                    ),
                                                ],
                                            ),
                                            dcc.Input(debounce=True,
                                                      id="exo-distance-input",
                                                      type="number",
                                                      min=0.5,
                                                      max=5.0,
                                                      step="any",
                                                      value=float(cfg.exo_gap),
                                                      style={"width": "100px", **ctrl_box_style},
                                                      ),
                                        ],
                                    ),
                                ],
                            ),

                            html.Div(
                                style={"marginTop": "10px", "boxSizing": "border-box"},
                                children=[
                                    html.Button(
                                        "Hide slice labels",
                                        id="btn-slice-toggle",
                                        n_clicks=0,
                                        style={"width": "100%", "whiteSpace": "nowrap"},
                                    ),
                                ],
                            ),

                            html.Div(
                                style={"display": "flex", "flexDirection": "row", "alignItems": "center", "gap": "10px",
                                       "marginTop": "10px", "width": "100%"},
                                children=[
                                    html.Label("Label font",
                                               style={**ctrl_label_style, "marginBottom": "0px", "width": "260px"}),
                                    html.Div(
                                        style=right_ctrl_single_style,
                                        children=[
                                            dcc.Dropdown(
                                                id="slice-font",
                                                options=[
                                                    {"label": "Arial", "value": "Arial"},
                                                    {"label": "Times New Roman", "value": "Times New Roman"},
                                                    {"label": "Helvetica", "value": "Helvetica"},
                                                    {"label": "Courier New", "value": "Courier New"},
                                                    {"label": "Georgia", "value": "Georgia"},
                                                ],
                                                value="Arial",
                                                clearable=False,
                                                optionHeight=dropdown_option_h,
                                                style={**ctrl_box_style, "width": "100%"},
                                            ),
                                        ],
                                    ),
                                ],
                            ),

                            html.Div(
                                style={"display": "flex", "flexDirection": "row", "alignItems": "center", "gap": "10px",
                                       "marginTop": "10px", "width": "100%"},
                                children=[
                                    html.Label("Label size",
                                               style={**ctrl_label_style, "marginBottom": "0px", "width": "260px"}),
                                    html.Div(
                                        style=right_ctrl_col_style,
                                        children=[
                                            html.Div(
                                                style={"flex": "1", "minWidth": "0"},
                                                children=[
                                                    dcc.Slider(
                                                        id="slice-size-slider",
                                                        min=4,
                                                        max=36,
                                                        step=1,
                                                        value=11,
                                                        updatemode="mouseup",
                                                        tooltip={"always_visible": False},
                                                    ),
                                                ],
                                            ),
                                            dcc.Input(debounce=True,
                                                      id="slice-size-input",
                                                      type="number",
                                                      min=6,
                                                      max=72,
                                                      step="any",
                                                      value=32,
                                                      style={"width": "100px", **ctrl_box_style},
                                                      ),
                                        ],
                                    ),
                                ],
                            ),

                            html.Div(
                                style={"display": "flex", "flexDirection": "row", "alignItems": "center", "gap": "10px",
                                       "marginTop": "10px", "width": "100%"},
                                children=[
                                    html.Label("Label color",
                                               style={**ctrl_label_style, "marginBottom": "0px", "width": "260px"}),
                                    html.Div(
                                        style=right_ctrl_col_style,
                                        children=[
                                            dcc.Input(debounce=True,
                                                      id="label-color-picker",
                                                      type="color",
                                                      value="#000000",
                                                      style={"width": "48px", "height": "48px"},
                                                      ),
                                            dcc.Input(
                                                id="label-color-text",
                                                type="text",
                                                value="#000000",
                                                debounce=True,
                                                style={"flex": "1", "minWidth": "0", **ctrl_box_style},
                                            ),
                                        ],
                                    ),
                                ],
                            ),

                            html.Div(
                                style={"display": "flex", "flexDirection": "row", "alignItems": "center", "gap": "10px",
                                       "marginTop": "10px", "width": "100%"},
                                children=[
                                    html.Label("Label distance",
                                               style={**ctrl_label_style, "marginBottom": "0px", "width": "260px"}),
                                    html.Div(
                                        style=right_ctrl_col_style,
                                        children=[
                                            html.Div(
                                                style={"flex": "1", "minWidth": "0"},
                                                children=[
                                                    dcc.Slider(
                                                        id="label-distance-slider",
                                                        min=0.0,
                                                        max=5.0,
                                                        step=0.1,
                                                        value=2.0,
                                                        updatemode="mouseup",
                                                        tooltip={"always_visible": False},
                                                    ),
                                                ],
                                            ),
                                            dcc.Input(debounce=True,
                                                      id="label-distance-input",
                                                      type="number",
                                                      min=0.0,
                                                      max=5.0,
                                                      step="any",
                                                      value=2.0,
                                                      style={"width": "100px", **ctrl_box_style},
                                                      ),
                                        ],
                                    ),
                                ],
                            ),

                            html.Div(
                                style={"display": "flex", "flexDirection": "row", "alignItems": "center", "gap": "10px",
                                       "marginTop": "10px", "width": "100%"},
                                children=[
                                    html.Label("Label rotation",
                                               style={**ctrl_label_style, "marginBottom": "0px", "width": "260px"}),
                                    html.Div(
                                        style=right_ctrl_col_style,
                                        children=[
                                            html.Div(
                                                style={"flex": "1", "minWidth": "0"},
                                                children=[
                                                    dcc.Slider(
                                                        id="slice-rotation-slider",
                                                        min=-180,
                                                        max=180,
                                                        step=1,
                                                        value=0,
                                                        updatemode="mouseup",
                                                        tooltip={"always_visible": False},
                                                    ),
                                                ],
                                            ),
                                            dcc.Input(debounce=True,
                                                      id="slice-rotation-input",
                                                      type="number",
                                                      min=-180,
                                                      max=180,
                                                      step="any",
                                                      value=0,
                                                      style={"width": "100px", **ctrl_box_style},
                                                      ),
                                        ],
                                    ),
                                ],
                            ),

                            # --- Slice Label ---
                            html.Hr(style={"width": "100%", "border": "none", "borderTop": "2px solid #CCCCCC",
                                           "margin": "14px 0 8px 0"}),
                            html.Label("Slice label", style={**ctrl_label_style, "marginBottom": "6px"}),
                            html.Div(id="slice-rename-container", children=[]),
                            # Collapse / Expand buttons row
                            html.Div(
                                id="collapse-btn-row",
                                style={"display": "flex", "flexDirection": "row", "gap": "10px", "marginTop": "12px"},
                                children=[
                                    html.Button("Collapse", id="btn-collapse", n_clicks=0,
                                                style={"flex": "1", "fontSize": "36px", "padding": "6px 0",
                                                       "cursor": "pointer",
                                                       "border": "2px solid #AAAAAA", "borderRadius": "4px",
                                                       "backgroundColor": "#F5F5F5"}),
                                    html.Button("Expand", id="btn-expand", n_clicks=0,
                                                style={"flex": "1", "fontSize": "36px", "padding": "6px 0",
                                                       "cursor": "pointer",
                                                       "border": "2px solid #AAAAAA", "borderRadius": "4px",
                                                       "backgroundColor": "#F5F5F5"}),
                                ],
                            ),
                            # Hide/Show collapse line button row
                            html.Div(
                                id="collapse-line-btn-row",
                                style={"display": "flex", "flexDirection": "row", "marginTop": "8px"},
                                children=[
                                    html.Button("Hide collapse line", id="btn-toggle-collapse-line", n_clicks=0,
                                                style={"flex": "1", "fontSize": "36px", "padding": "6px 0",
                                                       "cursor": "pointer",
                                                       "border": "2px solid #AAAAAA", "borderRadius": "4px",
                                                       "backgroundColor": "#F5F5F5"}),
                                ],
                            ),

                        ], style=panel_style),
                        # ========== BAND STYLE PANEL ==========
                        html.Div([
                            html.Div(
                                style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                                       "gap": "10px", "marginBottom": "10px"},
                                children=[
                                    html.Label("Band", style={**ctrl_label_style, "fontWeight": "700",
                                                              "marginBottom": "0px"}),
                                ],
                            ),

                            # Band mode switch button (global, not per-type)
                            html.Div(
                                style={"marginBottom": "10px", "boxSizing": "border-box"},
                                children=[
                                    html.Button(
                                        "Switch to Strength",
                                        id="btn-band-mode",
                                        n_clicks=0,
                                        style={"width": "100%"},
                                    ),
                                ],
                            ),

                            # Hide/Show all bands button
                            html.Div(
                                style={"marginBottom": "10px", "boxSizing": "border-box"},
                                children=[
                                    html.Button(
                                        "Show bands",
                                        id="btn-bands-all",
                                        n_clicks=0,
                                        style={"width": "100%"},
                                    ),
                                ],
                            ),

                            # Width / Proportion control (global, affects all types)
                            html.Div(
                                style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                                       "gap": "10px", "marginBottom": "10px"},
                                children=[
                                    html.Label("Width", id="band-width-label",
                                               style={**ctrl_label_style, "marginBottom": "0px",
                                                      "width": right_label_w}),
                                    html.Div(
                                        style=right_ctrl_col_style,
                                        children=[
                                            html.Div(
                                                style={"flex": "1", "minWidth": "0"},
                                                children=[
                                                    dcc.Slider(
                                                        id="band-width-slider",
                                                        min=0.02,
                                                        max=1.0,
                                                        step=0.01,
                                                        value=float(cfg.band_width_ratio),
                                                        updatemode="mouseup", tooltip={"always_visible": False},
                                                    ),
                                                ],
                                            ),
                                            dcc.Input(debounce=True,
                                                      id="band-width-input",
                                                      type="number",
                                                      min=0.02,
                                                      max=1.0,
                                                      step="any",
                                                      value=float(cfg.band_width_ratio),
                                                      style={"width": "100px", **ctrl_box_style},
                                                      ),
                                        ],
                                    ),
                                ],
                            ),

                            # Hidden dummy band-type store (replaces the old dropdown, keeps callbacks compatible)
                            dcc.Store(id="band-type", data="All"),

                            # ---- Per-type collapsible sections ----
                            html.Div(
                                style={"borderTop": "2px solid #E0E0E0", "paddingTop": "8px", "marginTop": "4px"},
                                children=[
                                    _build_band_type_section(bt, cfg, ctrl_label_style, ctrl_box_style,
                                                             right_ctrl_col_style, right_label_w)
                                    for bt in BAND_TYPE_UI_ORDER
                                ],
                            ),

                        ], style=panel_style),

                        # ========== BLOCK STYLE PANEL ==========
                        html.Div([
                            html.Div(
                                style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                                       "gap": "10px", "marginBottom": "10px"},
                                children=[
                                    html.Label("Block", style={**ctrl_label_style, "fontWeight": "700",
                                                               "marginBottom": "0px"}),
                                ],
                            ),

                            html.Div(
                                style={"display": "flex", "flexDirection": "column", "gap": "10px"},
                                children=[
                                    # Block mode switch button
                                    html.Div(
                                        style={"marginBottom": "5px"},
                                        children=[
                                            html.Button(
                                                "Switch to Strength",
                                                id="btn-block-mode",
                                                n_clicks=0,
                                                style={"width": "100%"},
                                            ),
                                        ],
                                    ),
                                    html.Div(
                                        style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                                               "justifyContent": "flex-start", "gap": "10px", "width": "100%"},
                                        children=[
                                            html.Label("Width", id="block-width-label",
                                                       style={**ctrl_label_style, "marginBottom": "0px",
                                                              "width": right_label_w}),
                                            html.Div(
                                                style=right_ctrl_col_style,
                                                children=[
                                                    html.Div(
                                                        style={"flex": "1", "minWidth": "0"},
                                                        children=[
                                                            dcc.Slider(
                                                                id="block-width-slider",
                                                                min=0.5,
                                                                max=5.0,
                                                                step=0.1,
                                                                value=float(cfg.block_width),
                                                                updatemode="mouseup", tooltip={"always_visible": False},
                                                            ),
                                                        ],
                                                    ),
                                                    dcc.Input(debounce=True,
                                                              id="block-width-input",
                                                              type="number",
                                                              min=0.5,
                                                              max=5.0,
                                                              step="any",
                                                              value=float(cfg.block_width),
                                                              style={"width": "100px", **ctrl_box_style},
                                                              ),
                                                ],
                                            ),
                                        ],
                                    ),

                                    html.Div(
                                        style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                                               "justifyContent": "flex-start", "gap": "10px", "width": "100%"},
                                        children=[
                                            html.Label("Height", style={**ctrl_label_style, "marginBottom": "0px",
                                                                        "width": right_label_w}),
                                            html.Div(
                                                style=right_ctrl_col_style,
                                                children=[
                                                    html.Div(
                                                        style={"flex": "1", "minWidth": "0"},
                                                        children=[
                                                            dcc.Slider(
                                                                id="block-height-slider",
                                                                min=0.5,
                                                                max=12.0,
                                                                step=0.1,
                                                                value=float(cfg.block_height),
                                                                updatemode="mouseup", tooltip={"always_visible": False},
                                                            ),
                                                        ],
                                                    ),
                                                    dcc.Input(debounce=True,
                                                              id="block-height-input",
                                                              type="number",
                                                              min=0.2,
                                                              max=20.0,
                                                              step="any",
                                                              value=float(cfg.block_height),
                                                              style={"width": "100px", **ctrl_box_style},
                                                              ),
                                                ],
                                            ),
                                        ],
                                    ),

                                    # --- Text padding (L/R row + T/B row, Enclosure Radius style) ---
                                    html.Div(
                                        style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                                               "justifyContent": "flex-start", "gap": "10px", "width": "100%"},
                                        children=[
                                            html.Label("Text padding", style={**ctrl_label_style, "marginBottom": "0px",
                                                                              "width": right_label_w}),
                                            html.Div(
                                                style=right_ctrl_single_style,
                                                children=[
                                                    html.Div(
                                                        style=pair_two_input_row_style,
                                                        children=[
                                                            html.Span("L", style=pair_small_label_style),
                                                            html.Div(),
                                                            dcc.Input(id="blk-text-pad-l", type="number", min=0,
                                                                      max=5.0, step="any", value=0, debounce=False,
                                                                      style=pair_input_style),
                                                            html.Div(),
                                                            html.Span("R", style=pair_small_label_style),
                                                            html.Div(),
                                                            dcc.Input(id="blk-text-pad-r", type="number", min=0,
                                                                      max=5.0, step="any", value=0, debounce=False,
                                                                      style=pair_input_style),
                                                        ],
                                                    ),
                                                ],
                                            ),
                                        ],
                                    ),
                                    html.Div(
                                        style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                                               "justifyContent": "flex-start", "gap": "10px", "width": "100%"},
                                        children=[
                                            html.Label("", style={**ctrl_label_style, "marginBottom": "0px",
                                                                  "width": right_label_w}),
                                            html.Div(
                                                style=right_ctrl_single_style,
                                                children=[
                                                    html.Div(
                                                        style=pair_two_input_row_style,
                                                        children=[
                                                            html.Span("T", style=pair_small_label_style),
                                                            html.Div(),
                                                            dcc.Input(id="blk-text-pad-t", type="number", min=0,
                                                                      max=5.0, step="any", value=0, debounce=False,
                                                                      style=pair_input_style),
                                                            html.Div(),
                                                            html.Span("B", style=pair_small_label_style),
                                                            html.Div(),
                                                            dcc.Input(id="blk-text-pad-b", type="number", min=0,
                                                                      max=5.0, step="any", value=0, debounce=False,
                                                                      style=pair_input_style),
                                                        ],
                                                    ),
                                                ],
                                            ),
                                        ],
                                    ),

                                    # Hidden dummy block-type store (replaces old dropdown)
                                    dcc.Store(id="block-type", data="All"),

                                    # ---- Per-type collapsible sections ----
                                    html.Div(
                                        style={"borderTop": "2px solid #E0E0E0", "paddingTop": "8px",
                                               "marginTop": "4px"},
                                        children=[
                                            _build_block_all_section(cfg, ctrl_label_style, ctrl_box_style)
                                            if bt == "All" else
                                            _build_block_type_section(bt, cfg, ctrl_label_style, ctrl_box_style)
                                            for bt in BLOCK_TYPE_UI_ORDER
                                        ],
                                    ),

                                    # Hidden dummy elements for old IDs (backward compat)
                                    dcc.Input(id="block-fill-color-picker", type="hidden", value="#CCCCCC"),
                                    dcc.Input(id="block-fill-color-text", type="hidden", value="#CCCCCC"),
                                    dcc.Input(id="block-fill-opacity-slider", type="hidden", value="1.0"),
                                    dcc.Input(id="block-fill-opacity-input", type="hidden", value="1.0"),
                                    dcc.Input(id="block-border-width-slider", type="hidden", value="1"),
                                    dcc.Input(id="block-border-width-input", type="hidden", value="1"),
                                    dcc.Input(id="block-line-style", type="hidden", value="solid"),
                                    dcc.Input(id="block-border-color-picker", type="hidden", value="#222222"),
                                    dcc.Input(id="block-border-color-text", type="hidden", value="#222222"),
                                    dcc.Input(id="block-border-opacity-slider", type="hidden", value="1.0"),
                                    dcc.Input(id="block-border-opacity-input", type="hidden", value="1.0"),
                                    dcc.Input(id="block-radius-tl", type="hidden", value="0"),
                                    dcc.Input(id="block-radius-tr", type="hidden", value="0"),
                                    dcc.Input(id="block-radius-bl", type="hidden", value="0"),
                                    dcc.Input(id="block-radius-br", type="hidden", value="0"),
                                    dcc.Input(id="block-text-font", type="hidden", value="Arial"),
                                    dcc.Input(id="block-text-size-slider", type="hidden", value="12"),
                                    dcc.Input(id="block-text-size-input", type="hidden", value="12"),
                                    dcc.Input(id="block-text-color-picker", type="hidden", value="#111111"),
                                    dcc.Input(id="block-text-color-text", type="hidden", value="#111111"),
                                    dcc.Input(id="block-text-align", type="hidden", value="center"),
                                    dcc.Input(id="block-line-spacing-slider", type="hidden", value="0"),
                                    dcc.Input(id="block-line-spacing", type="hidden", value="0"),
                                    dcc.Input(id="block-text-rotation-slider", type="hidden", value="90"),
                                    dcc.Input(id="block-text-rotation", type="hidden", value="90"),
                                ],
                            ),
                        ], style={**panel_style, "zIndex": 300}),

                        # ========== AGGREGATION ENCLOSURE STYLE PANEL ==========
                        html.Div([
                            html.Label("Aggregation",
                                       style={**ctrl_label_style, "fontWeight": "700", "marginBottom": "10px"}),
                            html.Div(
                                style={"display": "flex", "flexDirection": "column", "gap": "10px"},
                                children=[
                                    html.Div(
                                        style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                                               "justifyContent": "flex-start", "gap": "10px", "width": "100%"},
                                        children=[
                                            html.Label("Radius", style={**ctrl_label_style, "marginBottom": "0px",
                                                                        "width": right_label_w}),
                                            html.Div(
                                                style=right_ctrl_single_style,
                                                children=[
                                                    html.Div(
                                                        style=pair_two_input_row_style,
                                                        children=[
                                                            html.Span("TL", style=pair_small_label_style),
                                                            html.Div(),
                                                            dcc.Input(id="group-radius-tl", type="number", min=0,
                                                                      max=50, step="any", value=0, debounce=False,
                                                                      style=pair_input_style),
                                                            html.Div(),
                                                            html.Span("TR", style=pair_small_label_style),
                                                            html.Div(),
                                                            dcc.Input(id="group-radius-tr", type="number", min=0,
                                                                      max=50, step="any", value=0, debounce=False,
                                                                      style=pair_input_style),
                                                        ],
                                                    ),
                                                ],
                                            ),
                                        ],
                                    ),

                                    html.Div(
                                        style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                                               "justifyContent": "flex-start", "gap": "10px", "width": "100%"},
                                        children=[
                                            html.Label("", style={**ctrl_label_style, "marginBottom": "0px",
                                                                  "width": right_label_w}),
                                            html.Div(
                                                style=right_ctrl_single_style,
                                                children=[
                                                    html.Div(
                                                        style=pair_two_input_row_style,
                                                        children=[
                                                            html.Span("BL", style=pair_small_label_style),
                                                            html.Div(),
                                                            dcc.Input(id="group-radius-bl", type="number", min=0,
                                                                      max=50, step="any", value=0, debounce=False,
                                                                      style=pair_input_style),
                                                            html.Div(),
                                                            html.Span("BR", style=pair_small_label_style),
                                                            html.Div(),
                                                            dcc.Input(id="group-radius-br", type="number", min=0,
                                                                      max=50, step="any", value=0, debounce=False,
                                                                      style=pair_input_style),
                                                        ],
                                                    ),
                                                ],
                                            ),
                                        ],
                                    ),

                                    html.Div(
                                        style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                                               "justifyContent": "flex-start", "gap": "10px", "width": "100%"},
                                        children=[
                                            html.Label("Width", style={**ctrl_label_style, "marginBottom": "0px",
                                                                       "width": right_label_w}),
                                            html.Div(
                                                style=right_ctrl_col_style,
                                                children=[
                                                    html.Div(
                                                        style={"flex": "1", "minWidth": "0"},
                                                        children=[
                                                            dcc.Slider(
                                                                id="group-line-width-slider",
                                                                min=0.0,
                                                                max=4.0,
                                                                step=0.05,
                                                                value=float(cfg.group_enclosure_line_width),
                                                                updatemode="mouseup", tooltip={"always_visible": False},
                                                            ),
                                                        ],
                                                    ),
                                                    dcc.Input(debounce=True,
                                                              id="group-line-width-input",
                                                              type="number",
                                                              min=0.5,
                                                              max=10.0,
                                                              step="any",
                                                              value=float(cfg.group_enclosure_line_width),
                                                              style={"width": "100px", **ctrl_box_style},
                                                              ),
                                                ],
                                            ),
                                        ],
                                    ),

                                    html.Div(
                                        style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                                               "justifyContent": "flex-start", "gap": "10px", "width": "100%"},
                                        children=[
                                            html.Label("Line style", style={**ctrl_label_style, "marginBottom": "0px",
                                                                            "width": right_label_w}),
                                            html.Div(
                                                style=right_ctrl_single_style,
                                                children=[
                                                    dcc.Dropdown(
                                                        id="group-line-style",
                                                        options=[
                                                            {"label": "Solid", "value": "solid"},
                                                            {"label": "Dash", "value": "dash"},
                                                            {"label": "Dot", "value": "dot"},
                                                            {"label": "Dashdot", "value": "dashdot"},
                                                        ],
                                                        value="solid",
                                                        clearable=False,
                                                        optionHeight=dropdown_option_h,
                                                        style={**ctrl_box_style, "width": "100%"},
                                                    ),
                                                ],
                                            ),
                                        ],
                                    ),

                                    html.Div(
                                        style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                                               "justifyContent": "flex-start", "gap": "10px", "width": "100%"},
                                        children=[
                                            html.Label("Color", style={**ctrl_label_style, "marginBottom": "0px",
                                                                       "width": right_label_w}),
                                            html.Div(
                                                style=right_ctrl_col_style,
                                                children=[
                                                    dcc.Input(debounce=True, id="group-color-picker", type="color",
                                                              value="#000000",
                                                              style={"width": "48px", "height": "48px"}),
                                                    dcc.Input(id="group-color-text", type="text", value="#000000",
                                                              debounce=True,
                                                              style={"flex": "1", "minWidth": "0", **ctrl_box_style}),
                                                ],
                                            ),
                                        ],
                                    ),

                                    html.Div(
                                        style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                                               "justifyContent": "flex-start", "gap": "10px", "width": "100%"},
                                        children=[
                                            html.Label("Opacity", style={**ctrl_label_style, "marginBottom": "0px",
                                                                         "width": right_label_w}),
                                            html.Div(
                                                style=right_ctrl_col_style,
                                                children=[
                                                    html.Div(
                                                        style={"flex": "1", "minWidth": "0"},
                                                        children=[
                                                            dcc.Slider(
                                                                id="group-opacity-slider",
                                                                min=0.0,
                                                                max=1.0,
                                                                step=0.01,
                                                                value=float(cfg.group_enclosure_opacity),
                                                                updatemode="mouseup", tooltip={"always_visible": False},
                                                            ),
                                                        ],
                                                    ),
                                                    dcc.Input(debounce=True,
                                                              id="group-opacity-input",
                                                              type="number",
                                                              min=0.0,
                                                              max=1.0,
                                                              step="any",
                                                              value=float(cfg.group_enclosure_opacity),
                                                              style={"width": "100px", **ctrl_box_style},
                                                              ),
                                                ],
                                            ),
                                        ],
                                    ),

                                    html.Div(
                                        style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                                               "justifyContent": "flex-start", "gap": "10px", "width": "100%"},
                                        children=[
                                            html.Label("Inner gap", style={**ctrl_label_style, "marginBottom": "0px",
                                                                           "width": right_label_w}),
                                            html.Div(
                                                style=right_ctrl_col_style,
                                                children=[
                                                    html.Div(
                                                        style={"flex": "1", "minWidth": "0"},
                                                        children=[
                                                            dcc.Slider(
                                                                id="group-inner-gap-slider",
                                                                min=0.0,
                                                                max=1.0,
                                                                step=0.01,
                                                                value=float(cfg.inner_gap),
                                                                updatemode="mouseup", tooltip={"always_visible": False},
                                                            ),
                                                        ],
                                                    ),
                                                    dcc.Input(debounce=True,
                                                              id="group-inner-gap-input",
                                                              type="number",
                                                              min=0.0,
                                                              max=1.0,
                                                              step="any",
                                                              value=float(cfg.inner_gap),
                                                              style={"width": "100px", **ctrl_box_style},
                                                              ),
                                                ],
                                            ),
                                        ],
                                    ),

                                    html.Div(
                                        style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                                               "justifyContent": "flex-start", "gap": "10px", "width": "100%"},
                                        children=[
                                            html.Label("Outer gap", style={**ctrl_label_style, "marginBottom": "0px",
                                                                           "width": right_label_w}),
                                            html.Div(
                                                style=right_ctrl_col_style,
                                                children=[
                                                    html.Div(
                                                        style={"flex": "1", "minWidth": "0"},
                                                        children=[
                                                            dcc.Slider(
                                                                id="group-outer-gap-slider",
                                                                min=0.0,
                                                                max=4.0,
                                                                step=0.01,
                                                                value=float(cfg.outer_gap),
                                                                updatemode="mouseup", tooltip={"always_visible": False},
                                                            ),
                                                        ],
                                                    ),
                                                    dcc.Input(debounce=True,
                                                              id="group-outer-gap-input",
                                                              type="number",
                                                              min=0.0,
                                                              max=10.0,
                                                              step="any",
                                                              value=float(cfg.outer_gap),
                                                              style={"width": "100px", **ctrl_box_style},
                                                              ),
                                                ],
                                            ),
                                        ],
                                    ),

                                    html.Div(
                                        style={"boxSizing": "border-box"},
                                        children=[
                                            html.Button(
                                                "Hide aggregation name",
                                                id="btn-group-label-toggle",
                                                n_clicks=0,
                                                style={"width": "100%"},
                                            ),
                                        ],
                                    ),

                                    html.Div(
                                        style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                                               "justifyContent": "flex-start", "gap": "10px", "width": "100%"},
                                        children=[
                                            html.Label("Name font", style={**ctrl_label_style, "marginBottom": "0px",
                                                                           "width": right_label_w}),
                                            html.Div(
                                                style=right_ctrl_single_style,
                                                children=[
                                                    dcc.Dropdown(
                                                        id="group-label-font",
                                                        options=[
                                                            {"label": "Arial", "value": "Arial"},
                                                            {"label": "Times New Roman", "value": "Times New Roman"},
                                                            {"label": "Helvetica", "value": "Helvetica"},
                                                            {"label": "Courier New", "value": "Courier New"},
                                                            {"label": "Georgia", "value": "Georgia"},
                                                        ],
                                                        value="Arial",
                                                        clearable=False,
                                                        optionHeight=dropdown_option_h,
                                                        style={**ctrl_box_style, "width": "100%"},
                                                    ),
                                                ],
                                            ),
                                        ],
                                    ),

                                    html.Div(
                                        style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                                               "justifyContent": "flex-start", "gap": "10px", "width": "100%"},
                                        children=[
                                            html.Label("Name size", style={**ctrl_label_style, "marginBottom": "0px",
                                                                           "width": right_label_w}),
                                            html.Div(
                                                style=right_ctrl_col_style,
                                                children=[
                                                    html.Div(
                                                        style={"flex": "1", "minWidth": "0"},
                                                        children=[
                                                            dcc.Slider(
                                                                id="group-label-size-slider",
                                                                min=4,
                                                                max=24,
                                                                step=1,
                                                                value=8,
                                                                updatemode="mouseup", tooltip={"always_visible": False},
                                                            ),
                                                        ],
                                                    ),
                                                    dcc.Input(debounce=True,
                                                              id="group-label-size-input",
                                                              type="number",
                                                              min=6,
                                                              max=48,
                                                              step="any",
                                                              value=32,
                                                              style={"width": "100px", **ctrl_box_style},
                                                              ),
                                                ],
                                            ),
                                        ],
                                    ),

                                    html.Div(
                                        style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                                               "justifyContent": "flex-start", "gap": "10px", "width": "100%"},
                                        children=[
                                            html.Label("Name color", style={**ctrl_label_style, "marginBottom": "0px",
                                                                            "width": right_label_w}),
                                            html.Div(
                                                style=right_ctrl_col_style,
                                                children=[
                                                    dcc.Input(debounce=True, id="group-label-color-picker",
                                                              type="color", value="#000000",
                                                              style={"width": "48px", "height": "48px"}),
                                                    dcc.Input(id="group-label-color-text", type="text", value="#000000",
                                                              debounce=True,
                                                              style={"flex": "1", "minWidth": "0", **ctrl_box_style}),
                                                ],
                                            ),
                                        ],
                                    ),

                                    html.Div(
                                        style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                                               "justifyContent": "flex-start", "gap": "10px", "width": "100%"},
                                        children=[
                                            html.Label("Distance", style={**ctrl_label_style, "marginBottom": "0px",
                                                                          "width": right_label_w}),
                                            html.Div(
                                                style=right_ctrl_col_style,
                                                children=[
                                                    html.Div(
                                                        style={"flex": "1", "minWidth": "0"},
                                                        children=[
                                                            dcc.Slider(
                                                                id="group-label-offset-x-slider",
                                                                min=0.0,
                                                                max=2.0,
                                                                step=0.01,
                                                                value=float(cfg.enclosure_group_label_offset_x),
                                                                updatemode="mouseup", tooltip={"always_visible": False},
                                                            ),
                                                        ],
                                                    ),
                                                    dcc.Input(debounce=True,
                                                              id="group-label-offset-x-input",
                                                              type="number",
                                                              min=0.0,
                                                              max=2.0,
                                                              step="any",
                                                              value=float(cfg.enclosure_group_label_offset_x),
                                                              style={"width": "100px", **ctrl_box_style},
                                                              ),
                                                ],
                                            ),
                                        ],
                                    ),

                                    html.Div(
                                        style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                                               "justifyContent": "flex-start", "gap": "10px", "width": "100%"},
                                        children=[
                                            html.Label("Name rotation",
                                                       style={**ctrl_label_style, "marginBottom": "0px",
                                                              "width": right_label_w}),
                                            html.Div(
                                                style=right_ctrl_col_style,
                                                children=[
                                                    html.Div(
                                                        style={"flex": "1", "minWidth": "0"},
                                                        children=[
                                                            dcc.Slider(
                                                                id="group-text-rotation-slider",
                                                                min=-180,
                                                                max=180,
                                                                step=1,
                                                                value=90,
                                                                updatemode="mouseup",
                                                                tooltip={"always_visible": False},
                                                            ),
                                                        ],
                                                    ),
                                                    dcc.Input(debounce=True,
                                                              id="group-text-rotation",
                                                              type="number",
                                                              min=-180,
                                                              max=180,
                                                              step="any",
                                                              value=90,
                                                              style={"width": "100px", **ctrl_box_style},
                                                              ),
                                                ],
                                            ),
                                        ],
                                    ),

                                    # --- Line spacing (aggregation name) ---
                                    html.Div(
                                        style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                                               "justifyContent": "flex-start", "gap": "10px", "width": "100%"},
                                        children=[
                                            html.Label("Line spacing", style={**ctrl_label_style, "marginBottom": "0px",
                                                                              "width": right_label_w}),
                                            html.Div(
                                                style=right_ctrl_col_style,
                                                children=[
                                                    html.Div(
                                                        style={"flex": "1", "minWidth": "0"},
                                                        children=[
                                                            dcc.Slider(
                                                                id="group-label-line-spacing-slider",
                                                                min=-10,
                                                                max=60,
                                                                step=1,
                                                                value=0,
                                                                updatemode="mouseup",
                                                                tooltip={"always_visible": False},
                                                            ),
                                                        ],
                                                    ),
                                                    dcc.Input(debounce=True,
                                                              id="group-label-line-spacing-input",
                                                              type="number",
                                                              min=-20,
                                                              max=200,
                                                              step="any",
                                                              value=0,
                                                              style={"width": "100px", **ctrl_box_style},
                                                              ),
                                                ],
                                            ),
                                        ],
                                    ),

                                    # --- Padding (T / B) ---
                                    html.Div(
                                        style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                                               "justifyContent": "flex-start", "gap": "10px", "width": "100%"},
                                        children=[
                                            html.Label("Padding", style={**ctrl_label_style, "marginBottom": "0px",
                                                                         "width": right_label_w}),
                                            html.Div(
                                                style=right_ctrl_single_style,
                                                children=[
                                                    html.Div(
                                                        style=pair_two_input_row_style,
                                                        children=[
                                                            html.Span("T", style=pair_small_label_style),
                                                            html.Div(),
                                                            dcc.Input(id="group-label-pad-t", type="number", min=0,
                                                                      max=5.0, step="any", value=0, debounce=False,
                                                                      style=pair_input_style),
                                                            html.Div(),
                                                            html.Span("B", style=pair_small_label_style),
                                                            html.Div(),
                                                            dcc.Input(id="group-label-pad-b", type="number", min=0,
                                                                      max=5.0, step="any", value=0, debounce=False,
                                                                      style=pair_input_style),
                                                        ],
                                                    ),
                                                ],
                                            ),
                                        ],
                                    ),
                                ],
                            ),
                        ], style={**panel_style, "zIndex": 200}),

                        # ========== SUPERGROUP ENCLOSURE STYLE PANELS (DYNAMIC) ==========
                        html.Div(id="supergroup-panels-container", children=[], style={"marginBottom": "10px"}),

                    ],
                ),
                html.Div(id="right-panel-spacer", style={"height": "0px", "flexShrink": "0"}),
            ],
        ),

        dcc.Store(id="store-base-fig", data=fig0.to_dict()),
        dcc.Store(id="store-meta", data=meta0),
        dcc.Store(id="store-ilog-event", data=None),  # interaction log event sink
        dcc.Store(id="store-selected", data={"selected_id": None}),
        dcc.Store(id="store-exo", data={"hidden_slices": [], "nonce": {}, "hide_all_initial": True}),
        dcc.Store(id="store-click-nonce", data={"nonce": {}}),
        dcc.Store(id="store-zoom-state", data=None),
        dcc.Store(id="store-level-labels", data=level_state0),
        dcc.Store(id="store-enclosure-visibility", data={}),
        # {"1": True, "2": True, ...} - enclosure border visibility
        dcc.Store(id="store-bands", data={"hidden_types": ["Inflow", "Outflow"]}),
        dcc.Store(id="store-band-mode", data={"mode": "existence", "proportion": 0.5}),  # "existence" or "strength"
        # Per-type on/off state: {"Outflow": {"color_on": False, "opacity_on": False}, ...}
        dcc.Store(id="store-band-type-onoff",
                  data={bt: {"color_on": False, "opacity_on": False} for bt in BAND_TYPE_UI_ORDER}),
        # Hidden dummy elements for backward-compatible callbacks (band-color-picker, band-color-text, band-opacity-slider, band-opacity-input)
        html.Div(style={"display": "none"}, children=[
            dcc.Input(id="band-color-picker", type="hidden", value="#999999"),
            dcc.Input(id="band-color-text", type="hidden", value="#999999"),
            dcc.Slider(id="band-opacity-slider", min=0, max=1, step=0.01, value=float(cfg.band_default_opacity)),
            dcc.Input(id="band-opacity-input", type="hidden", value=float(cfg.band_default_opacity)),
        ]),
        dcc.Store(id="store-block-mode", data={"mode": "existence", "median_width": float(cfg.block_width)}),
        # "existence" or "strength"
        dcc.Store(
            id="store-style",
            data={
                "band": {
                    "colors": dict(cfg.band_colors),
                    "opacities": {k: float(cfg.band_default_opacity) for k in cfg.band_colors.keys()},
                    "width_ratios": {k: float(cfg.band_width_ratio) for k in cfg.band_colors.keys()},
                    "per_type_overrides": {bt: {"color_on": False, "opacity_on": False,
                                                "color": str(cfg.band_colors.get(bt, "#999999")),
                                                "opacity": float(cfg.band_default_opacity)}
                                           for bt in BAND_TYPE_UI_ORDER},
                },

                "block": {
                    "widths": {k: float(cfg.block_width) for k in cfg.block_colors.keys()},
                    "heights": {k: float(cfg.block_height) for k in cfg.block_colors.keys()},
                    "fill_colors": dict(cfg.block_colors),
                    "fill_opacities": {k: 1.0 for k in cfg.block_colors.keys()},
                    "border_colors": {k: "#222222" for k in cfg.block_colors.keys()},
                    "border_opacities": {k: 1.0 for k in cfg.block_colors.keys()},
                    "border_radii": {k: [0, 0, 0, 0] for k in cfg.block_colors.keys()},
                    "line_styles": {k: "solid" for k in cfg.block_colors.keys()},
                    "text_fonts": {k: "Arial" for k in cfg.block_colors.keys()},
                    "text_sizes": {k: int(cfg.block_text_size) for k in cfg.block_colors.keys()},
                    "text_colors": dict(cfg.block_text_colors),
                    "text_aligns": {k: "center" for k in cfg.block_colors.keys()},
                    "text_rotations": {k: 90 for k in cfg.block_colors.keys()},
                    "border_widths": {k: float(cfg.block_border_width) for k in cfg.block_colors.keys()},
                },

                "group": {
                    "line_width": float(cfg.group_enclosure_line_width),
                    "opacity": float(cfg.group_enclosure_opacity),
                    "color": "#000000",
                    "radii": [0, 0, 0, 0],
                    "line_style": "solid",
                    "label_font": "Arial",
                    "label_size": 8,
                    "label_color": "#000000",
                    "label_offset_x": float(cfg.enclosure_group_label_offset_x),
                    "label_rotation": 90,
                    "label_line_spacing": 0,
                    "label_pad_b": 0.0,
                    "label_pad_t": 0.0,
                },

                "supergroups": supergroups_style0,

                "layout": {
                    "layer_gap": float(cfg.layer_gap),
                    "exo_gap": float(cfg.exo_gap),
                    "inner_gap": float(cfg.inner_gap),
                    "outer_gap": float(cfg.outer_gap),
                },
            },
        ),
        dcc.Store(id="store-slice-style",
                  data={"visible": True, "font": "Arial", "size": 11, "rotation": 0, "color": "#000000",
                        "distance": 2.0}),
        dcc.Store(id="store-slice-rename", data={}),
        dcc.Store(id="store-collapse",
                  data={"anchor_x": None, "anchor_y": None, "collapsed": False, "line_hidden": False}),
        # Separate trigger for rebuild_base — only increments on actual Collapse/Expand/Toggle-line,
        # NOT on checkbox anchor selection (which must not trigger a figure rebuild).
        dcc.Store(id="store-collapse-trigger", data=0),
        # Multi-layer collapse state
        dcc.Store(id="store-upload-slices", data={"paths": []}),
        dcc.Store(id="store-dl-trigger", data={"t": None}),
        dcc.Store(id="store-graph-size", data={"w": None, "h": None}),
        dcc.Store(id="store-sweep-params", data={"k_max": 10, "m": 2, "delta": 0.01}),
        # Remembers the user-dragged position of each movable popover so that
        # closing + reopening restores the last dragged location instead of
        # snapping back to the button anchor. Keyed by popover element id.
        # Populated by the drag handler via `popover-drag-trigger` below.
        dcc.Store(id="store-popover-positions", data={}),
        # Hidden trigger button used by the JS drag handler to flush the latest
        # position map into `store-popover-positions`. The handler writes JSON
        # onto this button's `data-payload` dataset attribute and then clicks it.
        html.Button(id="popover-drag-trigger", n_clicks=0, style={"display": "none"}),
        # Loading overlay flags — one per heavy callback; overlay shown when any is True
        dcc.Store(id="store-loading-upload", data=False),
        dcc.Store(id="store-loading-generate", data=False),
        dcc.Store(id="store-loading-rebuild", data=False),
        # Click-to-highlight (and other render_with_highlight triggers).
        # This callback handles the user-visible response to block/band clicks
        # and is the most user-perceptible source of "did my click register?"
        # latency. The flag flips True the moment Dash queues the callback on
        # the client (sub-frame latency), giving the user immediate feedback.
        dcc.Store(id="store-loading-highlight", data=False),
        # One-shot interval to measure real canvas size before first user interaction
        dcc.Interval(id="init-size-interval", interval=200, n_intervals=0, max_intervals=3),

        # ── File-order modal (multi-file mode) ─────────────────────────
        html.Div(
            id="file-order-modal",
            children=[
                html.Div(
                    id="file-order-dialog",
                    children=[
                        html.Div(
                            style={"display": "flex", "justifyContent": "space-between",
                                   "alignItems": "flex-start", "marginBottom": "6px"},
                            children=[
                                html.H3("Arrange Slices"),
                                html.Span(
                                    "×",
                                    id="btn-order-modal-close",
                                    n_clicks=0,
                                    style={"fontSize": "44px", "lineHeight": "44px", "cursor": "pointer",
                                           "userSelect": "none", "color": "#888", "padding": "0 4px"},
                                ),
                            ],
                        ),
                        html.Div(
                            "Use ↑ ↓ to reorder the uploaded files, then click Generate.",
                            className="subtitle",
                        ),
                        html.Div(id="file-order-upload-hint"),
                        html.Div(id="file-order-list"),
                        html.Div(id="file-order-error",
                                 style={"fontSize": "30px", "color": "#CC0000",
                                        "minHeight": "28px", "marginTop": "6px"}),
                        html.Div(
                            id="file-order-actions",
                            children=[
                                html.Button("Cancel", id="btn-order-cancel", n_clicks=0),
                                html.Button("Generate", id="btn-order-generate", n_clicks=0),
                            ],
                        ),
                    ],
                ),
            ],
        ),
    ], style={"height": "100vh", "overflow": "hidden", "position": "relative"})

    # --- Loading overlay + top-bar status: driven by FOUR flag stores ---
    # Both the canvas overlay and the top-bar "Updating" badge are toggled
    # by the same JS function. The badge text reflects which action is in
    # progress: Uploading / Generating / Updating.
    #
    # The badge ALSO bridges the gap between "server callback finishes" and
    # "Plotly has actually drawn the figure". The server-side `running=`
    # parameter only flips the store back to False when the Python callback
    # returns; it does NOT account for network transmission of the figure
    # JSON or for Plotly's parse/layout/draw time on the client (which can
    # take 1-3 seconds for complex figures). Without bridging this gap, the
    # user sees: badge → blank → figure, which is confusing.
    #
    # The bridge: when server flag transitions true→false, set a
    # `pending_render` flag and show "Rendering…". The flag is cleared by
    # Plotly's `plotly_afterplot` event (handled in the figure-change
    # clientside callback below). A 30s safety timeout clears it if afterplot
    # fails to fire (e.g., callback returned no_update for graph.figure).
    #
    # The transition false→true is debounced by 50ms to absorb the brief gap
    # between two consecutive server callbacks (e.g., rebuild_base → render_
    # with_highlight chained via store-base-fig). Without this, the badge
    # would briefly flicker to "Rendering…" between them.
    app.clientside_callback(
        """
function(upload, generate, rebuild, highlight) {
    // ----- One-time bootstrap: shared state + update helper on `window` -----
    if (!window._birdcage_status_state) {
        window._birdcage_status_state = {
            server_active: false,
            pending_render: false,
            last_server_label: 'Updating\u2026',
        };
        window._birdcage_update_status_badge = function() {
            var st = window._birdcage_status_state;
            var active = st.server_active || st.pending_render;
            var badge = document.getElementById('updating-status');
            if (!badge) return;
            badge.style.display = active ? 'flex' : 'none';
            var lbl = badge.querySelector('.updating-label');
            if (lbl) {
                if (st.server_active) {
                    lbl.textContent = st.last_server_label;
                } else if (st.pending_render) {
                    lbl.textContent = 'Rendering\u2026';
                }
            }
        };
    }

    // ----- Compute new server state -----
    var prev_server = window._birdcage_status_state.server_active;
    var server_active = !!(upload || generate || rebuild || highlight);
    window._birdcage_status_state.server_active = server_active;

    if (server_active) {
        var lbl_txt = 'Updating\u2026';
        if (upload)        lbl_txt = 'Uploading\u2026';
        else if (generate) lbl_txt = 'Generating\u2026';
        else if (rebuild)  lbl_txt = 'Updating\u2026';
        else if (highlight) lbl_txt = 'Updating\u2026';
        window._birdcage_status_state.last_server_label = lbl_txt;
    }

    // ----- Server transition true→false (debounced 50ms) -----
    if (prev_server && !server_active) {
        if (window._birdcage_false_transition_timer) {
            clearTimeout(window._birdcage_false_transition_timer);
        }
        window._birdcage_false_transition_timer = setTimeout(function() {
            if (window._birdcage_status_state.server_active) return;
            window._birdcage_status_state.pending_render = true;
            if (window._birdcage_pending_render_timer) {
                clearTimeout(window._birdcage_pending_render_timer);
            }
            window._birdcage_pending_render_timer = setTimeout(function() {
                window._birdcage_status_state.pending_render = false;
                window._birdcage_pending_render_timer = null;
                window._birdcage_update_status_badge();
            }, 30000);
            window._birdcage_update_status_badge();
        }, 50);
    } else if (!prev_server && server_active) {
        if (window._birdcage_false_transition_timer) {
            clearTimeout(window._birdcage_false_transition_timer);
            window._birdcage_false_transition_timer = null;
        }
    }

    // ----- Canvas-level loading overlay (existing behavior) -----
    var el = document.getElementById('graph-loading-overlay');
    if (el) {
        el.className = server_active ? 'loading-overlay active' : 'loading-overlay';
    }

    // ----- Top-bar badge -----
    window._birdcage_update_status_badge();
    return window.dash_clientside.no_update;
}
""",
        Output("graph-loading-overlay", "className"),
        Input("store-loading-upload", "data"),
        Input("store-loading-generate", "data"),
        Input("store-loading-rebuild", "data"),
        Input("store-loading-highlight", "data"),
        prevent_initial_call=True,
    )

    # --- Popover drag sync: when the JS drag handler clicks the hidden
    # `popover-drag-trigger` button, read the JSON payload it staged on its
    # `data-payload` dataset attribute and write it into
    # `store-popover-positions`. The toggle_*_popover callbacks read this
    # store so that reopening a popover restores its last dragged position.
    app.clientside_callback(
        """
function(n) {
    if (!n) return window.dash_clientside.no_update;
    var btn = document.getElementById('popover-drag-trigger');
    if (!btn) return window.dash_clientside.no_update;
    var raw = btn.dataset.payload || '';
    if (!raw) return window.dash_clientside.no_update;
    try {
        return JSON.parse(raw);
    } catch (e) {
        return window.dash_clientside.no_update;
    }
}
""",
        Output("store-popover-positions", "data"),
        Input("popover-drag-trigger", "n_clicks"),
        prevent_initial_call=True,
    )

    # --- Force Plotly to recompute its size after figure updates (fix initial off-center until user zoom/scroll) ---
    app.clientside_callback(
        """
        function(fig) {
            if (fig) {
                setTimeout(function(){
                    window.dispatchEvent(new Event('resize'));
                }, 60);

                // Arm the gate so the next plotly_afterplot is recognized
                // as "render of this figure done". Hover-induced afterplots
                // that fire BEFORE the figure prop changed do not have this
                // flag and are ignored.
                window._birdcage_expecting_afterplot = true;
            }
            return null;
        }
        """,
        Output("__resize_trigger", "data"),
        Input("graph", "figure"),
        prevent_initial_call=True,
    )

    # --- Fix B: Measure canvas size once at page load (before any user interaction)
    # so the first rebuild_base uses real dimensions, not the 600×700 default.
    #
    # Cross-browser note: Chrome / Firefox / Safari / Edge can report
    # window.innerWidth / innerHeight values that differ by 10–30px for the
    # same visual window size (scrollbar handling, title-bar / menu-bar
    # heights, browser zoom level). Downstream, this would produce different
    # text-wrap decisions in Python, making the same figure look meaningfully
    # different across browsers. We quantize w/h to a 25px grid at every
    # write site so that small measurement variations collapse to the same
    # value, and the server-side text-fit sees a stable input. ---
    app.clientside_callback(
        """
function(n) {
    if (!n) return window.dash_clientside.no_update;
    function Q(v) { return Math.max(200, Math.round(v / 25) * 25); }
    // Reproduce uiScale formula from initCanvasControls (must match exactly)
    var _NAT_W = 920;
    var uiSc = Math.min(2.0, Math.max(0.15, window.innerWidth / (6 * _NAT_W)));
    // Canvas width = viewport minus VISUAL width of the right panel (layout_w × uiScale)
    var w = window.innerWidth - _NAT_W * uiSc;
    // Canvas height = viewport minus VISUAL height of the top bar
    var topBar = document.getElementById('ui-top-wrap');
    // getBoundingClientRect returns CSS-transformed visual size
    var topH = topBar ? Math.round(topBar.getBoundingClientRect().height) : Math.round(88 * uiSc);
    if (topH <= 0) topH = Math.round(88 * uiSc);
    var ht = window.innerHeight - topH;
    return {w: Q(w), h: Q(ht)};
}
""",
        Output("store-graph-size", "data", allow_duplicate=True),
        Input("init-size-interval", "n_intervals"),
        prevent_initial_call=True,
    )

    # --- Re-measure canvas size after every figure render.
    # CRITICAL: this MUST use the same viewport-derived formula as Path A
    # (init-size-interval) above. Earlier this read `gd._fullLayout.width`
    # and `getBoundingClientRect()`, which Firefox and Chrome interpret
    # differently in early figure-render frames — Chrome would commit a
    # smaller width (plotly's internal pre-layout value) for a tick,
    # writing it into store-graph-size and clobbering Path A's correct
    # value. Python then received a too-small _CANVAS_W, scaled
    # _ppu_x/_ppu_y down, and wrapped block text into 1-2 character
    # fragments with "..." ellipsis. Firefox didn't show the bug because
    # its `gd._fullLayout.width` matched Path A's value at the same point
    # in the render cycle.
    #
    # By using window.innerWidth / window.innerHeight here too — the
    # exact formula the top bar / right panel sizing uses (uiScale based
    # on `innerWidth / (6 * 920)`) — both paths produce identical values
    # in any browser at the same viewport size. The canvas dimensions
    # are now tied to the same source-of-truth as the rest of the UI,
    # which is exactly what makes the toolbar and right panel render
    # identically across browsers in the screenshots: same formula,
    # same result. ---
    app.clientside_callback(
        """
function(fig) {
    if (!fig) return window.dash_clientside.no_update;
    function Q(v) { return Math.max(200, Math.round(v / 25) * 25); }
    // Reproduce the EXACT viewport-derived formula from
    // init-size-interval above. Any divergence here would re-introduce
    // the cross-browser drift this callback exists to fix.
    var _NAT_W = 920;
    var uiSc = Math.min(2.0, Math.max(0.15, window.innerWidth / (6 * _NAT_W)));
    var w = window.innerWidth - _NAT_W * uiSc;
    var topBar = document.getElementById('ui-top-wrap');
    var topH = topBar ? Math.round(topBar.getBoundingClientRect().height) : Math.round(88 * uiSc);
    if (topH <= 0) topH = Math.round(88 * uiSc);
    var ht = window.innerHeight - topH;
    if (w > 0 && ht > 0) {
        return {w: Q(w), h: Q(ht), dpr: window.devicePixelRatio || 1.0};
    }
    return window.dash_clientside.no_update;
}
""",
        Output("store-graph-size", "data", allow_duplicate=True),
        Input("graph", "figure"),
        prevent_initial_call=True,
    )


    # Save y-axis range from user pan/zoom into store-zoom-state.
    # x-axis is preserved by uirevision. y-axis needs explicit save/restore
    # because scaleanchor="x" causes Plotly to recompute it on every figure update.
    # Also compute zoom_scale = initial_x_span / current_x_span so that
    # render_with_highlight can scale textfont.size and line.width accordingly.
    @app.callback(
        Output("store-zoom-state", "data"),
        Input("graph", "relayoutData"),
        State("store-zoom-state", "data"),
        State("store-meta", "data"),
        prevent_initial_call=True,
    )
    def save_y_range(relayout_data, current_zoom, meta):
        if not relayout_data:
            raise PreventUpdate
        current_zoom = current_zoom or {}
        updated = False
        if "yaxis.range[0]" in relayout_data and "yaxis.range[1]" in relayout_data:
            current_zoom["y"] = [relayout_data["yaxis.range[0]"], relayout_data["yaxis.range[1]"]]
            updated = True
        elif "yaxis.range" in relayout_data:
            current_zoom["y"] = list(relayout_data["yaxis.range"])
            updated = True
        if "xaxis.range[0]" in relayout_data and "xaxis.range[1]" in relayout_data:
            current_zoom["x"] = [relayout_data["xaxis.range[0]"], relayout_data["xaxis.range[1]"]]
            updated = True
        elif "xaxis.range" in relayout_data:
            current_zoom["x"] = list(relayout_data["xaxis.range"])
            updated = True
        # Compute zoom scale INCREMENTALLY: new_scale = old_scale * (prev_span / new_span).
        # This is level-independent: the scale only tracks HOW MUCH the user zoomed,
        # not the absolute data coordinates. Switching levels does not change cur_span
        # (uirevision preserves the viewport), so scale is automatically preserved
        # across level switches. Each scroll click produces the same scale increment
        # on every level.
        #
        # Bootstrap (first ever zoom): no prev_x_span saved yet, so derive the
        # starting scale from the current level's init_span so that scale=1.0 at the
        # default full view of whatever level is active first.
        #
        # The loop risk (render -> relayoutData -> render) is naturally broken:
        # after render, uirevision keeps the axis range unchanged, so Plotly
        # does NOT fire relayoutData again -> updated stays False -> PreventUpdate.
        x_now = current_zoom.get("x")
        if x_now and len(x_now) == 2:
            new_span = abs(float(x_now[1]) - float(x_now[0]))
            if new_span > 0:
                prev_span = current_zoom.get("prev_x_span")
                if prev_span and prev_span > 0:
                    # Incremental: scale by the ratio of viewport change only.
                    old_scale = float(current_zoom.get("scale", 1.0))
                    current_zoom["scale"] = old_scale * (prev_span / new_span)
                else:
                    # Bootstrap: first zoom ever — derive from current level's init_span
                    # so scale starts at 1.0 when the user is at the default full view.
                    init_xr = (meta or {}).get("initial_x_range")
                    if init_xr and len(init_xr) == 2:
                        init_span = abs(float(init_xr[1]) - float(init_xr[0]))
                        current_zoom["scale"] = init_span / new_span
                current_zoom["prev_x_span"] = new_span
        if not updated:
            raise PreventUpdate
        return current_zoom

    # ── Measure canvas at upload time (single-file path, no modal) ────────────
    app.clientside_callback(
        """
function(contents) {
    if (!contents) return window.dash_clientside.no_update;
    function Q(v) { return Math.max(200, Math.round(v / 25) * 25); }
    var _NAT_W = 920;
    var uiSc = Math.min(2.0, Math.max(0.15, window.innerWidth / (6 * _NAT_W)));
    var gd = document.getElementById('graph');
    var w = null, h = null;
    if (gd && gd._fullLayout) { w = gd._fullLayout.width || null; h = gd._fullLayout.height || null; }
    if (!w || w <= 0) w = window.innerWidth - _NAT_W * uiSc;
    if (!h || h <= 0) {
        var topBar = document.getElementById('ui-top-wrap');
        var topH = topBar ? Math.round(topBar.getBoundingClientRect().height) : Math.round(88 * uiSc);
        h = window.innerHeight - topH;
    }
    return {w: Q(w), h: Q(h), dpr: window.devicePixelRatio || 1.0};
}
""",
        Output("store-graph-size", "data", allow_duplicate=True),
        Input("upload-slices", "contents"),
        prevent_initial_call=True,
    )

    # ── Upload callback ───────────────────────────────────────────────
    @app.callback(
        Output("store-upload-slices", "data"),
        Output("upload-status", "children"),
        Output("store-pending-files", "data"),
        Output("store-zoom-state", "data", allow_duplicate=True),
        Input("upload-slices", "contents"),
        State("upload-slices", "filename"),
        prevent_initial_call=True,
        running=[
            (Output("upload-slices", "disabled"), True, False),
            (Output("store-loading-upload", "data"), True, False),
        ],
    )
    def validate_uploaded_slices(contents, filenames):
        if not contents or not filenames:
            raise PreventUpdate
        multi_mode = len(filenames) > 1
        try:
            if multi_mode:
                pending = _save_files_pending(contents, filenames)
                return dash.no_update, html.Span(
                    f"✓ {len(filenames)} file(s) ready — arrange order and click Generate.",
                    style={"color": "#1a73e8", "fontWeight": "500"},
                ), pending, dash.no_update   # zoom not reset yet — wait for generate
            else:
                result, status = _validate_uploaded_slices_inner(contents, filenames)
                return result, status, dash.no_update, {}   # reset zoom for new data
        except Exception as e:
            return dash.no_update, html.Span(
                f"⚠ Upload failed: {e}",
                style={"color": "#CC0000", "fontWeight": "600"},
            ), dash.no_update, dash.no_update

    def _save_files_pending(contents, filenames) -> dict:
        """Save uploaded files to temp dir and return pending-files store payload."""
        up_dir = Path(tempfile.gettempdir()) / "birdcage_uploads"
        up_dir.mkdir(parents=True, exist_ok=True)
        batch = uuid.uuid4().hex[:10]
        saved_paths: List[str] = []
        display_names: List[str] = []
        for fn, c in zip(filenames, contents):
            if not isinstance(c, str) or "," not in c:
                raise ValueError(f"Upload payload for '{fn}' is not a valid Dash upload string.")
            _header, b64 = c.split(",", 1)
            raw = base64.b64decode(b64)
            out_path = up_dir / f"{batch}__{Path(fn).name}"
            out_path.write_bytes(raw)
            # Validate it can be read at all
            _ = read_slice_excel(str(out_path))
            saved_paths.append(str(out_path))
            display_names.append(str(Path(fn).stem))
        return {"paths": saved_paths, "names": display_names}

    def _validate_uploaded_slices_inner(contents, filenames):
        # ── 解码第一个文件，写到临时目录 ──────────────────────────
        up_dir = Path(tempfile.gettempdir()) / "birdcage_uploads"
        up_dir.mkdir(parents=True, exist_ok=True)
        batch = uuid.uuid4().hex[:10]

        saved_paths: List[str] = []
        slice_names_local: List[str] = []

        # ── 判断模式：单文件多sheet 还是 多文件单sheet ──────────
        is_multisheet = (
                len(filenames) == 1 and
                Path(filenames[0]).suffix.lower() in [".xlsx", ".xls", ".ods"]
        )

        if is_multisheet:
            # 单文件多sheet模式
            fn = filenames[0]
            c = contents[0]
            if not isinstance(c, str) or "," not in c:
                raise ValueError(f"Upload payload for '{fn}' is not a valid Dash upload string.")
            _header, b64 = c.split(",", 1)
            raw = base64.b64decode(b64)
            out_path = up_dir / f"{batch}__{Path(fn).name}"
            out_path.write_bytes(raw)

            dfs_sheets, sheet_names = read_all_sheets(str(out_path))
            for sheet_df, sheet_name in zip(dfs_sheets, sheet_names):
                sheet_path = up_dir / f"{batch}__{sheet_name}.xlsx"
                sheet_df.to_excel(str(sheet_path), index=False)
                saved_paths.append(str(sheet_path))
                slice_names_local.append(sheet_name)
        else:
            # 多文件模式（单sheet each，按上传顺序）
            for fn, c in zip(filenames, contents):
                if not isinstance(c, str) or "," not in c:
                    raise ValueError(f"Upload payload for '{fn}' is not a valid Dash upload string.")
                _header, b64 = c.split(",", 1)
                raw = base64.b64decode(b64)
                out_path = up_dir / f"{batch}__{Path(fn).name}"
                out_path.write_bytes(raw)
                _ = read_slice_excel(str(out_path))
                saved_paths.append(str(out_path))
                slice_names_local.append(str(Path(fn).stem))

        # Compute category candidates
        cols0 = list(read_slice_excel(saved_paths[0]).columns)
        common_cols = set(cols0)
        for p in saved_paths[1:]:
            common_cols = common_cols.intersection(set(list(read_slice_excel(p).columns)))

        element_col_local = cols0[-1]
        category_candidates_local = [c for c in cols0[:-1] if (c != element_col_local and c in common_cols)]
        if not category_candidates_local:
            raise ValueError("No common category candidates found across uploaded slices.")

        default_category_local = category_candidates_local[-1]
        display_names_local = slice_names_local[:]

        return {
            "paths": saved_paths,
            "slice_names": slice_names_local,
            "display_names": display_names_local,
            "category_candidates": category_candidates_local,
            "default_category": default_category_local,
        }, ""

    # ── Show/hide file-order modal ────────────────────────────────────
    @app.callback(
        Output("file-order-modal", "className"),
        Input("store-pending-files", "data"),
        Input("btn-order-cancel", "n_clicks"),
        Input("btn-order-modal-close", "n_clicks"),
        prevent_initial_call=True,
    )
    def toggle_order_modal(pending, n_cancel, n_close):
        ctx = dash.callback_context
        trig = ctx.triggered[0]["prop_id"].split(".")[0] if ctx.triggered else ""
        if trig in ("btn-order-cancel", "btn-order-modal-close"):
            return ""
        # Close when pending files are cleared (= successful generate)
        if pending is not None and not pending.get("paths"):
            return ""
        # Show when pending files are present
        if pending and pending.get("paths"):
            return "visible"
        return ""

    # Modal dialog now uses pure CSS vw/em sizing — no JS scaling needed on open.

    # ── Cancel: reset pending store + clear status message ───────────
    @app.callback(
        Output("store-pending-files", "data", allow_duplicate=True),
        Output("upload-status", "children", allow_duplicate=True),
        Output("file-order-error", "children", allow_duplicate=True),
        Input("btn-order-cancel", "n_clicks"),
        Input("btn-order-modal-close", "n_clicks"),
        prevent_initial_call=True,
    )
    def cancel_order_dialog(n_cancel, n_close):
        ctx = dash.callback_context
        if not ctx.triggered or not ctx.triggered[0]["value"]:
            raise PreventUpdate
        return {"paths": [], "names": []}, "", ""

    # ── Render the draggable order list ───────────────────────────────
    @app.callback(
        Output("file-order-list", "children"),
        Output("file-order-upload-hint", "children"),
        Input("store-pending-files", "data"),
        prevent_initial_call=True,
    )
    def render_order_list(pending):
        if not pending or not pending.get("paths"):
            return [], ""
        paths = list(pending.get("paths", []))
        names = list(pending.get("names", []))
        n = len(paths)
        rows = []
        for i, name in enumerate(names):
            up_disabled = (i == 0)
            down_disabled = (i == n - 1)
            rows.append(
                html.Div(
                    className="order-row",
                    children=[
                        html.Span(f"{i + 1}.", className="order-seq"),
                        html.Span(name, className="order-name", title=name),
                        html.Button(
                            "↑",
                            id={"type": "file-move", "index": i, "dir": "up"},
                            n_clicks=0,
                            className="order-arrow-btn",
                            disabled=up_disabled,
                        ),
                        html.Button(
                            "↓",
                            id={"type": "file-move", "index": i, "dir": "down"},
                            n_clicks=0,
                            className="order-arrow-btn",
                            disabled=down_disabled,
                        ),
                    ],
                )
            )
        return rows, ""

    # ── Move file up/down in pending store ────────────────────────────
    @app.callback(
        Output("store-pending-files", "data", allow_duplicate=True),
        Input({"type": "file-move", "index": ALL, "dir": ALL}, "n_clicks"),
        State("store-pending-files", "data"),
        prevent_initial_call=True,
    )
    def reorder_pending_files(n_clicks_list, pending):
        ctx = dash.callback_context
        if not ctx.triggered or not any(v for v in (n_clicks_list or [])):
            raise PreventUpdate
        trig_str = ctx.triggered[0]["prop_id"]
        # Extract index and dir from pattern-match id
        try:
            btn_id = json.loads(trig_str.split(".")[0])
        except Exception:
            raise PreventUpdate
        if not ctx.triggered[0]["value"]:
            raise PreventUpdate
        idx = int(btn_id["index"])
        direction = btn_id["dir"]
        paths = list(pending.get("paths", []))
        names = list(pending.get("names", []))
        if direction == "up" and idx > 0:
            paths[idx - 1], paths[idx] = paths[idx], paths[idx - 1]
            names[idx - 1], names[idx] = names[idx], names[idx - 1]
        elif direction == "down" and idx < len(paths) - 1:
            paths[idx], paths[idx + 1] = paths[idx + 1], paths[idx]
            names[idx], names[idx + 1] = names[idx + 1], names[idx]
        else:
            raise PreventUpdate
        return {"paths": paths, "names": names}

    # ── Measure canvas size at the moment the user clicks Generate ───────────
    # This runs client-side immediately on click, before rebuild_base fires,
    # so the Python callback always gets accurate dimensions (not the old estimate).
    app.clientside_callback(
        """
function(n) {
    if (!n) return window.dash_clientside.no_update;
    function Q(v) { return Math.max(200, Math.round(v / 25) * 25); }
    var _NAT_W = 920;
    var uiSc = Math.min(2.0, Math.max(0.15, window.innerWidth / (6 * _NAT_W)));
    // Prefer actual Plotly dimensions if available
    var gd = document.getElementById('graph');
    var w = null, h = null;
    if (gd && gd._fullLayout) {
        w = gd._fullLayout.width  || null;
        h = gd._fullLayout.height || null;
    }
    if (!w || w <= 0) w = window.innerWidth - _NAT_W * uiSc;
    if (!h || h <= 0) {
        var topBar = document.getElementById('ui-top-wrap');
        var topH = topBar ? Math.round(topBar.getBoundingClientRect().height) : Math.round(88 * uiSc);
        h = window.innerHeight - topH;
    }
    return {w: Q(w), h: Q(h), dpr: window.devicePixelRatio || 1.0};
}
""",
        Output("store-graph-size", "data", allow_duplicate=True),
        Input("btn-order-generate", "n_clicks"),
        prevent_initial_call=True,
    )

    # ── Generate: validate ordered files → store-upload-slices ───────
    @app.callback(
        Output("store-upload-slices", "data", allow_duplicate=True),
        Output("upload-status", "children", allow_duplicate=True),
        Output("store-pending-files", "data", allow_duplicate=True),
        Output("file-order-error", "children"),
        Output("store-zoom-state", "data", allow_duplicate=True),
        Input("btn-order-generate", "n_clicks"),
        State("store-pending-files", "data"),
        prevent_initial_call=True,
        running=[
            (Output("btn-order-generate", "disabled"), True, False),
            (Output("store-loading-generate", "data"), True, False),
        ],
    )
    def generate_from_ordered_files(n_clicks, pending):
        if not n_clicks or not pending or not pending.get("paths"):
            raise PreventUpdate
        paths = list(pending.get("paths", []))
        names = list(pending.get("names", []))
        try:
            cols0 = list(read_slice_excel(paths[0]).columns)
            common_cols = set(cols0)
            for p in paths[1:]:
                common_cols = common_cols.intersection(set(list(read_slice_excel(p).columns)))
            element_col_local = cols0[-1]
            category_candidates_local = [c for c in cols0[:-1] if (c != element_col_local and c in common_cols)]
            if not category_candidates_local:
                raise ValueError("No common category columns found across uploaded files.")
            default_category_local = category_candidates_local[-1]
            result = {
                "paths": paths,
                "slice_names": names,
                "display_names": names,
                "category_candidates": category_candidates_local,
                "default_category": default_category_local,
            }
            # Success: clear pending + reset zoom for the new data
            return result, "", {"paths": [], "names": []}, "", {}
        except Exception as e:
            return dash.no_update, dash.no_update, dash.no_update, f"⚠ {e}", dash.no_update

    @app.callback(
        Output("file-details-popover", "style"),
        Output("file-details-popover-text", "children"),
        Input("btn-file-details", "n_clicks"),
        Input("btn-file-details-close", "n_clicks"),
        State("store-upload-slices", "data"),
        State("file-details-popover", "style"),
        State("store-popover-positions", "data"),
        prevent_initial_call=True,
    )
    def toggle_file_details_popover(n_open, n_close, upload_state, cur_style, popover_positions):
        ctx = dash.callback_context
        trig = ""
        if ctx and ctx.triggered:
            trig = str(ctx.triggered[0].get("prop_id", "")).split(".", 1)[0]

        base = dict(cur_style or {})
        base.setdefault("position", "absolute")
        base.setdefault("top", "100%")
        base.setdefault("left", "0px")
        base.setdefault("marginTop", "6px")
        base.setdefault("zIndex", "30000")
        base.setdefault("backgroundColor", "#FFFFFF")
        base.setdefault("border", "2px solid #CCCCCC")
        base.setdefault("borderRadius", "8px")
        base.setdefault("padding", "10px 12px")
        base.setdefault("minWidth", "360px")
        base.setdefault("boxShadow", "0 6px 18px rgba(0,0,0,0.18)")
        base.setdefault("whiteSpace", "pre-line")

        # If the user has previously dragged this popover, honour that position
        # instead of the anchor defaults above. Clear marginTop when an absolute
        # top has been set so we don't compound the anchor offset.
        _pos = (popover_positions or {}).get("file-details-popover") or {}
        if _pos.get("left"):
            base["left"] = _pos["left"]
        if _pos.get("top"):
            base["top"] = _pos["top"]
            base["marginTop"] = "0px"

        if trig == "btn-file-details-close":
            base["display"] = "none"
            return base, ""

        if base.get("display") == "block":
            base["display"] = "none"
            return base, ""

        if not upload_state or not upload_state.get("paths"):
            msg = "No files uploaded yet.\nClick 'Upload files' to add slice files."
            base["display"] = "block"
            return base, msg

        names = list(upload_state.get("display_names") or [])
        if not names:
            names = [str(Path(p).name).split("__", 1)[-1] for p in (upload_state.get("paths") or [])]

        msg = "Loaded " + str(len(names)) + " file(s):\n" + ",\n".join(names) + "."
        base["display"] = "block"
        return base, msg

    @app.callback(
        Output("guide-popover", "style"),
        Input("btn-guide", "n_clicks"),
        Input("btn-guide-close", "n_clicks"),
        State("guide-popover", "style"),
        State("store-popover-positions", "data"),
        prevent_initial_call=True,
    )
    def toggle_guide_popover(n_open, n_close, cur_style, popover_positions):
        ctx = dash.callback_context
        trig = ""
        if ctx and ctx.triggered:
            trig = str(ctx.triggered[0].get("prop_id", "")).split(".", 1)[0]

        base = dict(cur_style or {})
        base.setdefault("position", "absolute")
        base.setdefault("top", "100%")
        base.setdefault("left", "0px")
        base.setdefault("marginTop", "6px")
        base.setdefault("zIndex", "30000")
        base.setdefault("backgroundColor", "#FFFFFF")
        base.setdefault("border", "2px solid #CCCCCC")
        base.setdefault("borderRadius", "8px")
        base.setdefault("padding", "30px 30px")
        base.setdefault("minWidth", "440px")
        base.setdefault("boxShadow", "0 6px 18px rgba(0,0,0,0.18)")
        base.setdefault("whiteSpace", "pre-line")

        # Restore last dragged position if the user has moved this popover.
        _pos = (popover_positions or {}).get("guide-popover") or {}
        if _pos.get("left"):
            base["left"] = _pos["left"]
        if _pos.get("top"):
            base["top"] = _pos["top"]
            base["marginTop"] = "0px"

        if trig == "btn-guide-close":
            base["display"] = "none"
            return base

        if base.get("display") == "block":
            base["display"] = "none"
            return base

        base["display"] = "block"
        return base

    @app.callback(
        Output("category-col", "options"),
        Output("category-col", "value"),
        Output("category-col-wrap", "style"),
        Input("store-upload-slices", "data"),
        State("category-col", "value"),
        prevent_initial_call=True,
    )
    def update_category_options_from_upload(upload_state, current_value):
        if not upload_state or not upload_state.get("paths"):
            raise PreventUpdate

        cands = list(upload_state.get("category_candidates") or [])
        if not cands:
            raise PreventUpdate

        opts = [{"label": c, "value": c} for c in cands]
        v = current_value if (current_value in cands) else str(upload_state.get("default_category") or cands[-1])

        # Compute the longest option's BOLD-rendered width.
        # The selected option in the dropdown menu renders bold, so we need
        # to size for that. Bold Title-Case capital-heavy strings ("NAICS
        # Industry") are ~25-30% wider than the global k=0.60 estimate.
        # k=0.85 + 1.12 cushion is empirically safe.
        font_size = float(cfg.ui_font_size)
        max_chars = max((len(str(c)) for c in cands), default=1)
        longest_w = float(max_chars) * font_size * 0.85 * 1.12

        # chrome on the value side: arrow zone (54) + L padding (14) +
        # R padding (24) + outer border (2) + cushion (16)
        chrome_w = 110.0

        # Lock the WRAPPER DIV's width (not the outer combo's). Why:
        # react-select inside dcc.Dropdown sizes itself based on the current
        # value, which means anything inside the dropdown can change width
        # when the user picks a different level. By wrapping the Dropdown
        # in a fixed-width div, the OUTER container's width stays constant.
        # The combo's overall width = label box + this wrapper, which is
        # also constant, so the whole control never resizes.
        wrap_w = int(max(360.0, longest_w + chrome_w))
        _w_px = str(wrap_w) + "px"
        wrap_style = {
            "flex": "0 0 auto",         # don't grow/shrink
            "width": _w_px,
            "minWidth": _w_px,
            "maxWidth": _w_px,
            "alignSelf": "stretch",
        }

        return opts, v, wrap_style

    # Style panel (Band) callbacks

    def _clamp01(x: Any, default: float) -> float:
        # Handle None default
        if default is None:
            default = 0.0
        try:
            if x is None or x == "" or x == "None":
                return float(default)
            v = float(x)
        except (ValueError, TypeError):
            v = float(default)
        return float(max(0.0, min(1.0, v)))

    def _clamp_band_width(x: Any, default: float) -> float:
        # Handle None default
        if default is None:
            default = 0.10
        try:
            if x is None or x == "" or x == "None":
                return float(default)
            v = float(x)
        except (ValueError, TypeError):
            v = float(default)
        return float(max(0.02, min(0.30, v)))

    def _clamp_layer_gap(x: Any, default: float) -> float:
        # Handle None default
        if default is None:
            default = 1.0
        try:
            if x is None or x == "" or x == "None":
                return float(default)
            v = float(x)
        except (ValueError, TypeError):
            v = float(default)
        return float(max(1.0, min(20.0, v)))

    def _normalize_hex_color(s: Any, fallback: str = "#999999") -> str:
        if not isinstance(s, str):
            return fallback
        t = s.strip()
        if not t:
            return fallback
        if not t.startswith("#"):
            t = "#" + t
        if len(t) != 7:
            return fallback
        hexd = set("0123456789abcdefABCDEF")
        if any(ch not in hexd for ch in t[1:]):
            return fallback
        return "#" + t[1:].upper()

    @app.callback(
        Output("band-opacity-input", "value"),
        Output("band-opacity-slider", "value"),
        Input("band-type", "data"),
        prevent_initial_call=True,
    )
    def sync_band_opacity_stub(_band_type_data):
        # Stub: per-type opacity is now controlled by individual per-type controls
        raise PreventUpdate

    @app.callback(
        Output("layer-gap-input", "value"),
        Output("layer-gap-slider", "value"),
        Input("layer-gap-input", "value"),
        Input("layer-gap-slider", "value"),
        prevent_initial_call=True,
    )
    def sync_layer_gap(v_in, v_slider):
        ctx = dash.callback_context
        trig = (ctx.triggered[0]["prop_id"] if (ctx.triggered and len(ctx.triggered) > 0) else "")

        # Drop keys that are no longer valid (e.g., after changing dataset/category).

        if trig.startswith("layer-gap-input"):
            v = _clamp_layer_gap(v_in, float(cfg.layer_gap))
        else:
            v = _clamp_layer_gap(v_slider, float(cfg.layer_gap))
        return v, v

    def _clamp_exo_gap(x: Any, default: float) -> float:
        # Handle None default
        if default is None:
            default = 1.0
        try:
            if x is None or x == "" or x == "None":
                return float(default)
            v = float(x)
        except (ValueError, TypeError):
            v = float(default)
        return float(max(0.5, min(5.0, v)))

    @app.callback(
        Output("exo-distance-input", "value"),
        Output("exo-distance-slider", "value"),
        Input("exo-distance-input", "value"),
        Input("exo-distance-slider", "value"),
        prevent_initial_call=True,
    )
    def sync_exo_distance(v_in, v_slider):
        ctx = dash.callback_context
        trig = (ctx.triggered[0]["prop_id"] if (ctx.triggered and len(ctx.triggered) > 0) else "")

        if trig.startswith("exo-distance-input"):
            v = _clamp_exo_gap(v_in, float(cfg.exo_gap))
        else:
            v = _clamp_exo_gap(v_slider, float(cfg.exo_gap))
        return v, v

    def _clamp_slice_font_size(x: any, default: int) -> int:
        try:
            if x is None or x == "" or x == "None":
                return int(default)
            v = int(float(x))
        except (ValueError, TypeError):
            v = int(default)
        return int(max(6, min(72, v)))

    def _clamp_slice_rotation(x: any, default: int) -> int:
        try:
            if x is None or x == "" or x == "None":
                return int(default)
            v = int(float(x))
        except (ValueError, TypeError):
            v = int(default)
        return int(max(-180, min(180, v)))

    def _clamp_label_distance(x: any, default: float) -> float:
        try:
            if x is None or x == "" or x == "None":
                return float(default)
            v = float(x)
        except (ValueError, TypeError):
            v = float(default)
        return float(max(0.0, min(5.0, v)))

    @app.callback(
        Output("slice-size-input", "value"),
        Output("slice-size-slider", "value"),
        Input("slice-size-input", "value"),
        Input("slice-size-slider", "value"),
        prevent_initial_call=True,
    )
    def sync_slice_size(v_in, v_slider):
        ctx = dash.callback_context
        trig = (ctx.triggered[0]["prop_id"] if (ctx.triggered and len(ctx.triggered) > 0) else "")
        if trig.startswith("slice-size-input"):
            v = _clamp_slice_font_size(v_in, 11)
        else:
            v = _clamp_slice_font_size(v_slider, 11)
        return v, v

    @app.callback(
        Output("slice-rotation-input", "value"),
        Output("slice-rotation-slider", "value"),
        Input("slice-rotation-input", "value"),
        Input("slice-rotation-slider", "value"),
        prevent_initial_call=True,
    )
    def sync_slice_rotation(v_in, v_slider):
        ctx = dash.callback_context
        trig = (ctx.triggered[0]["prop_id"] if (ctx.triggered and len(ctx.triggered) > 0) else "")
        if trig.startswith("slice-rotation-input"):
            v = _clamp_slice_rotation(v_in, 0)
        else:
            v = _clamp_slice_rotation(v_slider, 0)
        return v, v

    @app.callback(
        Output("label-distance-input", "value"),
        Output("label-distance-slider", "value"),
        Input("label-distance-input", "value"),
        Input("label-distance-slider", "value"),
        prevent_initial_call=True,
    )
    def sync_label_distance(v_in, v_slider):
        ctx = dash.callback_context
        trig = (ctx.triggered[0]["prop_id"] if (ctx.triggered and len(ctx.triggered) > 0) else "")
        if trig.startswith("label-distance-input"):
            v = _clamp_label_distance(v_in, 0.0)
        else:
            v = _clamp_label_distance(v_slider, 0.0)
        return v, v

    @app.callback(
        Output("label-color-text", "value"),
        Output("label-color-picker", "value"),
        Input("label-color-text", "value"),
        Input("label-color-picker", "value"),
        State("store-slice-style", "data"),
        prevent_initial_call=True,
    )
    def sync_label_color(text_v, picker_v, slice_style_state):
        st = slice_style_state or {}
        cur = _normalize_hex_color(str(st.get("color", "#000000")), "#000000")

        ctx = dash.callback_context
        trig = (ctx.triggered[0]["prop_id"] if ctx.triggered else "")

        if trig.startswith("label-color-picker"):
            v = _normalize_hex_color(picker_v, cur)
            return v, v

        v = _normalize_hex_color(text_v, cur)
        return v, v

    @app.callback(
        Output("store-slice-style", "data"),
        Output("btn-slice-toggle", "children"),
        Input("btn-slice-toggle", "n_clicks"),
        Input("slice-font", "value"),
        Input("slice-size-slider", "value"),
        Input("slice-size-input", "value"),
        Input("slice-rotation-slider", "value"),
        Input("slice-rotation-input", "value"),
        Input("label-distance-slider", "value"),
        Input("label-distance-input", "value"),
        Input("label-color-text", "value"),
        Input("label-color-picker", "value"),
        State("store-slice-style", "data"),
        prevent_initial_call=True,
    )
    def update_slice_style(_n_clicks, font_v, size_slider_v, size_input_v, rot_slider_v, rot_input_v, dist_slider_v,
                           dist_input_v, color_text_v, color_picker_v, st):
        st = st or {"visible": True, "font": "Arial", "size": 11, "rotation": 0, "color": "#000000", "distance": 2.0}
        st = dict(st)

        ctx = dash.callback_context
        trig = (ctx.triggered[0]["prop_id"] if (ctx.triggered and len(ctx.triggered) > 0) else "")

        if trig.startswith("btn-slice-toggle"):
            st["visible"] = not bool(st.get("visible", True))

        if font_v is not None:
            st["font"] = str(font_v)

        # Use input value if input was triggered, otherwise use slider value
        if trig.startswith("slice-size-input"):
            st["size"] = _clamp_slice_font_size(size_input_v, int(st.get("size", 11)))
        else:
            st["size"] = _clamp_slice_font_size(size_slider_v, int(st.get("size", 11)))

        if trig.startswith("slice-rotation-input"):
            st["rotation"] = _clamp_slice_rotation(rot_input_v, int(st.get("rotation", 0)))
        else:
            st["rotation"] = _clamp_slice_rotation(rot_slider_v, int(st.get("rotation", 0)))

        if trig.startswith("label-distance-input"):
            st["distance"] = _clamp_label_distance(dist_input_v, float(st.get("distance", 2.0)))
        else:
            st["distance"] = _clamp_label_distance(dist_slider_v, float(st.get("distance", 2.0)))

        if trig.startswith("label-color-picker"):
            st["color"] = _normalize_hex_color(str(color_picker_v), str(st.get("color", "#000000")))
        else:
            st["color"] = _normalize_hex_color(str(color_text_v), str(st.get("color", "#000000")))

        btn_text = "Hide slice labels" if bool(st.get("visible", True)) else "Show slice labels"
        return st, btn_text

    @app.callback(
        Output("band-color-text", "value"),
        Output("band-color-picker", "value"),
        Input("band-type", "data"),
        prevent_initial_call=True,
    )
    def sync_band_color_stub(_band_type_data):
        # Stub: per-type color is now controlled by individual per-type controls
        raise PreventUpdate

    # --- Per-type band color/opacity: pattern-matching callbacks ---

    # 1) Toggle collapse of per-type sections
    @app.callback(
        Output({"type": "band-type-body", "index": ALL}, "style"),
        Output({"type": "band-type-arrow", "index": ALL}, "children"),
        Input({"type": "band-type-header", "index": ALL}, "n_clicks"),
        State({"type": "band-type-body", "index": ALL}, "style"),
        prevent_initial_call=True,
    )
    def toggle_band_type_sections(n_clicks_list, current_styles):
        ctx = dash.callback_context
        if not ctx.triggered:
            raise PreventUpdate

        # Determine which header was clicked
        trig_id = ctx.triggered[0]["prop_id"]
        # trig_id looks like '{"index":"Outflow","type":"band-type-header"}.n_clicks'
        try:
            trig_json = json.loads(trig_id.rsplit(".", 1)[0])
            clicked_index = trig_json["index"]
        except Exception:
            raise PreventUpdate

        out_styles = []
        out_arrows = []
        for i, bt in enumerate(BAND_TYPE_UI_ORDER):
            cur = dict(current_styles[i]) if current_styles[i] else {}
            if bt == clicked_index:
                # Toggle
                if cur.get("visibility") == "visible":
                    cur = {"visibility": "hidden", "height": "0", "overflow": "hidden", "padding": "8px 4px 8px 22px"}
                    out_arrows.append("▶")
                else:
                    cur = {"visibility": "visible", "height": "auto", "overflow": "visible",
                           "padding": "8px 4px 8px 22px"}
                    out_arrows.append("▼")
            else:
                out_arrows.append("▼" if cur.get("visibility") == "visible" else "▶")
            out_styles.append(cur)

        return out_styles, out_arrows

    # 2) On/Off toggle for color
    @app.callback(
        Output({"type": "band-color-onoff", "index": MATCH}, "children"),
        Output({"type": "band-color-onoff", "index": MATCH}, "style"),
        Output({"type": "band-color-picker-pt", "index": MATCH}, "disabled"),
        Output({"type": "band-color-text-pt", "index": MATCH}, "disabled"),
        Input({"type": "band-color-onoff", "index": MATCH}, "n_clicks"),
        prevent_initial_call=True,
    )
    def toggle_band_color_onoff(n_clicks):
        is_on = (n_clicks or 0) % 2 == 1
        label = "On" if is_on else "Off"
        style = {"minWidth": "48px", "fontSize": "28px", "padding": "2px 6px",
                 "borderRadius": "4px", "border": "2px solid #CCC",
                 "cursor": "pointer"}
        if is_on:
            style["backgroundColor"] = "#C8E6C9"
            style["color"] = "#2E7D32"
            style["fontWeight"] = "700"
        else:
            style["backgroundColor"] = "#F0F0F0"
            style["color"] = "#333"
            style["fontWeight"] = "400"
        return label, style, not is_on, not is_on

    # 3) On/Off toggle for opacity
    @app.callback(
        Output({"type": "band-opacity-onoff", "index": MATCH}, "children"),
        Output({"type": "band-opacity-onoff", "index": MATCH}, "style"),
        Output({"type": "band-opacity-slider-pt", "index": MATCH}, "disabled"),
        Output({"type": "band-opacity-input-pt", "index": MATCH}, "disabled"),
        Input({"type": "band-opacity-onoff", "index": MATCH}, "n_clicks"),
        prevent_initial_call=True,
    )
    def toggle_band_opacity_onoff(n_clicks):
        is_on = (n_clicks or 0) % 2 == 1
        label = "On" if is_on else "Off"
        style = {"minWidth": "48px", "fontSize": "28px", "padding": "2px 6px",
                 "borderRadius": "4px", "border": "2px solid #CCC",
                 "cursor": "pointer"}
        if is_on:
            style["backgroundColor"] = "#C8E6C9"
            style["color"] = "#2E7D32"
            style["fontWeight"] = "700"
        else:
            style["backgroundColor"] = "#F0F0F0"
            style["color"] = "#333"
            style["fontWeight"] = "400"
        return label, style, not is_on, not is_on

    # 4) Sync per-type opacity slider/input
    @app.callback(
        Output({"type": "band-opacity-input-pt", "index": MATCH}, "value"),
        Output({"type": "band-opacity-slider-pt", "index": MATCH}, "value"),
        Input({"type": "band-opacity-input-pt", "index": MATCH}, "value"),
        Input({"type": "band-opacity-slider-pt", "index": MATCH}, "value"),
        prevent_initial_call=True,
    )
    def sync_band_opacity_pt(v_in, v_slider):
        ctx = dash.callback_context
        trig = (ctx.triggered[0]["prop_id"] if ctx.triggered else "")
        if "input" in trig:
            v = _clamp01(v_in, cfg.band_default_opacity)
        else:
            v = _clamp01(v_slider, cfg.band_default_opacity)
        return v, v

    # 5) Sync per-type color picker/text
    @app.callback(
        Output({"type": "band-color-text-pt", "index": MATCH}, "value"),
        Output({"type": "band-color-picker-pt", "index": MATCH}, "value"),
        Output({"type": "band-type-swatch", "index": MATCH}, "style"),
        Input({"type": "band-color-text-pt", "index": MATCH}, "value"),
        Input({"type": "band-color-picker-pt", "index": MATCH}, "value"),
        State({"type": "band-color-onoff", "index": MATCH}, "n_clicks"),
        prevent_initial_call=True,
    )
    def sync_band_color_pt(text_v, picker_v, onoff_clicks):
        ctx = dash.callback_context
        trig = (ctx.triggered[0]["prop_id"] if ctx.triggered else "")
        # Determine the band type from the trigger
        try:
            trig_json = json.loads(trig.rsplit(".", 1)[0])
            bt = trig_json["index"]
        except Exception:
            bt = "Unknown"

        is_on = (onoff_clicks or 0) % 2 == 1
        default_color = str(cfg.band_colors.get(bt, "#999999"))

        if "picker" in trig:
            v = _normalize_hex_color(picker_v, default_color)
        else:
            v = _normalize_hex_color(text_v, default_color)

        # Update swatch color
        swatch_color = v
        swatch_style = {
            "width": "16px", "height": "16px", "borderRadius": "3px",
            "backgroundColor": swatch_color, "border": "2px solid #999",
            "marginRight": "8px", "flexShrink": "0",
        }

        return v, v, swatch_style

    # 6) Master update_style_store: merges per-type overrides + global width + layout gaps
    @app.callback(
        Output("store-style", "data"),
        Input({"type": "band-color-text-pt", "index": ALL}, "value"),
        Input({"type": "band-color-picker-pt", "index": ALL}, "value"),
        Input({"type": "band-color-onoff", "index": ALL}, "n_clicks"),
        Input({"type": "band-opacity-slider-pt", "index": ALL}, "value"),
        Input({"type": "band-opacity-onoff", "index": ALL}, "n_clicks"),
        Input("band-width-slider", "value"),
        Input("layer-gap-slider", "value"),
        Input("exo-distance-slider", "value"),
        State("store-style", "data"),
        State("store-band-mode", "data"),
        prevent_initial_call=True,
    )
    def update_style_store(color_texts, color_pickers, color_onoffs,
                           opacity_sliders, opacity_onoffs,
                           width_v, layer_gap_v, exo_gap_v,
                           style_state, band_mode_state):
        style_state = style_state or {}
        band_state = dict(style_state.get("band", {}) or {})

        colors = dict(band_state.get("colors") or {})
        op_map = dict(band_state.get("opacities") or {})
        w_map = dict(band_state.get("width_ratios") or {})
        per_type_overrides = dict(band_state.get("per_type_overrides") or {})

        for k in cfg.band_colors.keys():
            colors.setdefault(k, cfg.band_colors[k])
            op_map.setdefault(k, cfg.band_default_opacity)
            w_map.setdefault(k, cfg.band_width_ratio)

        # Apply per-type overrides from pattern-matching inputs
        # On/Off only controls whether the UI is disabled; values always come from controls.
        # Initially controls hold defaults. Once user turns On and edits, values persist even after Off.
        for i, bt in enumerate(BAND_TYPE_UI_ORDER):
            default_color = str(cfg.band_colors.get(bt, "#999999"))
            default_opacity = float(cfg.band_default_opacity)

            color_on = (color_onoffs[i] or 0) % 2 == 1 if i < len(color_onoffs) else False
            opacity_on = (opacity_onoffs[i] or 0) % 2 == 1 if i < len(opacity_onoffs) else False

            # Always use the current control value (not default) regardless of On/Off
            if i < len(color_texts):
                c = _normalize_hex_color(color_texts[i], default_color)
                colors[bt] = c
            else:
                colors[bt] = default_color

            if i < len(opacity_sliders):
                o = _clamp01(opacity_sliders[i], default_opacity)
                op_map[bt] = o
            else:
                op_map[bt] = default_opacity

            per_type_overrides[bt] = {
                "color_on": color_on,
                "opacity_on": opacity_on,
                "color": colors[bt],
                "opacity": op_map[bt],
            }

        # Global width (applies to all types equally)
        current_band_mode = (band_mode_state or {}).get("mode", "existence")
        if current_band_mode == "existence":
            v = _clamp_band_width(width_v, cfg.band_width_ratio)
            for k in w_map:
                w_map[k] = v

        band_state["colors"] = colors
        band_state["opacities"] = op_map
        band_state["width_ratios"] = w_map
        band_state["per_type_overrides"] = per_type_overrides
        style_state["band"] = band_state

        layout = dict(style_state.get("layout", {}) or {})
        layout["layer_gap"] = _clamp_layer_gap(layer_gap_v, cfg.layer_gap)
        layout["exo_gap"] = _clamp_exo_gap(exo_gap_v, cfg.exo_gap)
        style_state["layout"] = layout

        return style_state

    # 7) Per-type Hide/Show toggle → store-bands
    @app.callback(
        Output("store-bands", "data", allow_duplicate=True),
        Input({"type": "band-type-hide-btn", "index": ALL}, "n_clicks"),
        State("store-bands", "data"),
        prevent_initial_call=True,
    )
    def toggle_band_type_visibility(n_clicks_list, bands_state):
        ctx = dash.callback_context
        if not ctx.triggered:
            raise PreventUpdate

        trig_id = ctx.triggered[0]["prop_id"]
        try:
            trig_json = json.loads(trig_id.rsplit(".", 1)[0])
            clicked_bt = trig_json["index"]
        except Exception:
            raise PreventUpdate

        bands_state = bands_state or {"hidden_types": []}
        hidden_types = set(str(x) for x in (bands_state.get("hidden_types", []) or []))

        # If previously in __ALL__ mode, expand it to individual types first
        if "__ALL__" in hidden_types:
            hidden_types.discard("__ALL__")
            for bt in BAND_TYPE_UI_ORDER:
                hidden_types.add(bt)

        # Toggle the clicked type
        if clicked_bt in hidden_types:
            hidden_types.discard(clicked_bt)
        else:
            hidden_types.add(clicked_bt)

        return {"hidden_types": sorted(list(hidden_types))}

    # 8) Update per-type Hide/Show button labels based on store-bands
    @app.callback(
        Output({"type": "band-type-hide-btn", "index": ALL}, "children"),
        Output({"type": "band-type-hide-btn", "index": ALL}, "style"),
        Input("store-bands", "data"),
    )
    def update_band_hide_btn_labels(bands_state):
        hidden_types = set(str(x) for x in ((bands_state or {}).get("hidden_types", []) or []))
        hide_all = "__ALL__" in hidden_types

        base_style = {
            "fontSize": "24px", "padding": "2px 10px", "borderRadius": "4px",
            "border": "2px solid #CCC", "backgroundColor": "#FFF", "color": "#333",
            "cursor": "pointer", "width": "100px", "flexShrink": "0",
            "lineHeight": "1.4", "textAlign": "center",
        }

        labels = []
        styles = []
        for bt in BAND_TYPE_UI_ORDER:
            if hide_all or bt in hidden_types:
                labels.append("Show")
            else:
                labels.append("Hide")
            styles.append(dict(base_style))
        return labels, styles

        # Style panel (Block) callbacks(x: Any, default: float) -> float:
        # Handle None default
        if default is None:
            default = 1.0
        try:
            if x is None or x == "" or x == "None":
                return float(default)
            v = float(x)
        except (ValueError, TypeError):
            v = float(default)
        return float(max(0.5, min(5.0, v)))

    def _clamp_block_height(x: Any, default: float) -> float:
        # Handle None default
        if default is None:
            default = 1.0
        try:
            if x is None or x == "" or x == "None":
                return float(default)
            v = float(x)
        except (ValueError, TypeError):
            v = float(default)
        return float(max(0.2, min(20.0, v)))

    def _clamp_block_width(x: Any, default: float) -> float:
        if default is None:
            default = 1.0
        try:
            if x is None or x == "" or x == "None":
                return float(default)
            v = float(x)
        except (ValueError, TypeError):
            v = float(default)
        return float(max(0.5, min(5.0, v)))

    def _clamp_border_width(x: Any, default: float) -> float:
        # Handle None default
        if default is None:
            default = 1.0
        try:
            if x is None or x == "" or x == "None":
                return float(default)
            v = float(x)
        except (ValueError, TypeError):
            v = float(default)
        return float(max(0.0, min(5.0, v)))

    def _clamp_text_size(x: Any, default: int) -> int:
        # Handle None default
        if default is None:
            default = 12
        try:
            if x is None or x == "" or x == "None":
                return int(default)
            v = int(x)
        except (ValueError, TypeError):
            v = int(default)
        return int(max(8, min(72, v)))

    def _clamp_radius(x: Any, default: int = 0) -> int:
        if x is None or x == "" or x == "None":
            return int(default)
        try:
            v = int(float(x))  # Use float first to handle decimal strings
        except (ValueError, TypeError):
            v = int(default)
        return int(max(0, min(50, v)))

    def _safe_radii_list(raw_list: Any) -> List[int]:
        """Safely convert a list to a list of 4 integers for radii."""
        if not raw_list or not isinstance(raw_list, (list, tuple)) or len(raw_list) < 4:
            return [0, 0, 0, 0]
        result = []
        for x in raw_list[:4]:
            if x is None or x == '' or x == 'None':
                result.append(0)
            else:
                try:
                    result.append(int(float(x)))
                except (ValueError, TypeError):
                    result.append(0)
        return result

    # ============================================================
    # Block panel: per-type pattern-matching callbacks
    # ============================================================

    # Stub: old sync callbacks now unused (block-type is a hidden Store)
    @app.callback(
        Output("block-fill-color-text", "value"),
        Output("block-fill-color-picker", "value"),
        Input("block-type", "data"),
        prevent_initial_call=True,
    )
    def sync_block_fill_color_stub(_):
        raise PreventUpdate

    @app.callback(
        Output("block-border-color-text", "value"),
        Output("block-border-color-picker", "value"),
        Input("block-type", "data"),
        prevent_initial_call=True,
    )
    def sync_block_border_color_stub(_):
        raise PreventUpdate

    @app.callback(
        Output("block-text-color-text", "value"),
        Output("block-text-color-picker", "value"),
        Input("block-type", "data"),
        prevent_initial_call=True,
    )
    def sync_block_text_color_stub(_):
        raise PreventUpdate

    @app.callback(
        Output("block-fill-opacity-input", "value"),
        Output("block-fill-opacity-slider", "value"),
        Input("block-type", "data"),
        prevent_initial_call=True,
    )
    def sync_block_fill_opacity_stub(_):
        raise PreventUpdate

    @app.callback(
        Output("block-border-opacity-input", "value"),
        Output("block-border-opacity-slider", "value"),
        Input("block-type", "data"),
        prevent_initial_call=True,
    )
    def sync_block_border_opacity_stub(_):
        raise PreventUpdate

    @app.callback(
        Output("block-height-input", "value"),
        Output("block-height-slider", "value"),
        Input("block-type", "data"),
        Input("block-height-input", "value"),
        Input("block-height-slider", "value"),
        State("store-style", "data"),
        prevent_initial_call=True,
    )
    def sync_block_height(block_type_data, v_in, v_slider, style_state):
        ctx = dash.callback_context
        trig = (ctx.triggered[0]["prop_id"] if ctx.triggered else "")
        if trig.startswith("block-type"):
            raise PreventUpdate
        v = v_in if trig.startswith("block-height-input") else v_slider
        if v is None: v = float(cfg.block_height)
        v = float(max(0.2, min(20.0, float(v))))
        return v, v

    @app.callback(
        Output("block-border-width-input", "value"),
        Output("block-border-width-slider", "value"),
        Input("block-type", "data"),
        prevent_initial_call=True,
    )
    def sync_block_border_width_stub(_):
        raise PreventUpdate

    @app.callback(
        Output("block-text-size-input", "value"),
        Output("block-text-size-slider", "value"),
        Input("block-type", "data"),
        prevent_initial_call=True,
    )
    def sync_block_text_size_stub(_):
        raise PreventUpdate

    @app.callback(
        Output("block-line-spacing", "value"),
        Output("block-line-spacing-slider", "value"),
        Input("block-type", "data"),
        prevent_initial_call=True,
    )
    def sync_block_line_spacing_stub(_):
        raise PreventUpdate

    @app.callback(
        Output("block-text-rotation", "value"),
        Output("block-text-rotation-slider", "value"),
        Input("block-type", "data"),
        prevent_initial_call=True,
    )
    def sync_block_text_rotation_stub(_):
        raise PreventUpdate

    @app.callback(
        Output("block-radius-tl", "value"),
        Output("block-radius-tr", "value"),
        Output("block-radius-bl", "value"),
        Output("block-radius-br", "value"),
        Output("block-line-style", "value"),
        Output("block-text-font", "value"),
        Output("block-text-align", "value"),
        Input("block-type", "data"),
        prevent_initial_call=True,
    )
    def sync_block_other_params_stub(_):
        raise PreventUpdate

    # --- Block type section collapse/expand ---
    @app.callback(
        Output({"type": "blk-type-body", "index": ALL}, "style"),
        Output({"type": "blk-type-arrow", "index": ALL}, "children"),
        Input({"type": "blk-type-header", "index": ALL}, "n_clicks"),
        prevent_initial_call=True,
    )
    def toggle_block_type_sections(n_clicks_list):
        ctx = dash.callback_context
        if not ctx.triggered:
            raise PreventUpdate
        trig_id = ctx.triggered[0]["prop_id"]
        try:
            trig_json = json.loads(trig_id.rsplit(".", 1)[0])
            clicked_bt = trig_json["index"]
        except Exception:
            raise PreventUpdate
        body_styles = []
        arrows = []
        for i, bt in enumerate(BLOCK_TYPE_UI_ORDER):
            nc = n_clicks_list[i] or 0
            if bt == clicked_bt:
                is_open = nc % 2 == 1
            else:
                is_open = nc % 2 == 1
            body_styles.append({"visibility": "visible", "height": "auto", "overflow": "visible",
                                "padding": "8px 4px 8px 22px"} if is_open else {"visibility": "hidden", "height": "0",
                                                                                "overflow": "hidden",
                                                                                "padding": "8px 4px 8px 22px"})
            arrows.append("▼" if is_open else "▶")
        return body_styles, arrows

    # --- Generic On/Off toggles for each control group ---
    # Color-type On/Off (fill-color, border-color, text-color)
    _BLK_COLOR_SUFFIXES = ["fill-color", "border-color", "text-color"]

    @app.callback(
        Output({"type": "blk-fill-color-onoff", "index": MATCH}, "children"),
        Output({"type": "blk-fill-color-onoff", "index": MATCH}, "style"),
        Output({"type": "blk-fill-color-picker-pt", "index": MATCH}, "disabled"),
        Output({"type": "blk-fill-color-text-pt", "index": MATCH}, "disabled"),
        Input({"type": "blk-fill-color-onoff", "index": MATCH}, "n_clicks"),
        prevent_initial_call=True,
    )
    def toggle_blk_fill_color_onoff(nc):
        on = (nc or 0) % 2 == 1
        label = "On" if on else "Off"
        style = {"minWidth": "48px", "fontSize": "28px", "padding": "2px 6px",
                 "borderRadius": "4px", "border": "2px solid #CCC",
                 "cursor": "pointer"}
        if on:
            style["backgroundColor"] = "#C8E6C9"
            style["color"] = "#2E7D32"
            style["fontWeight"] = "700"
        else:
            style["backgroundColor"] = "#F0F0F0"
            style["color"] = "#333"
            style["fontWeight"] = "400"
        return label, style, not on, not on

    @app.callback(
        Output({"type": "blk-fill-color-text-pt", "index": MATCH}, "value"),
        Output({"type": "blk-fill-color-picker-pt", "index": MATCH}, "value"),
        Input({"type": "blk-fill-color-text-pt", "index": MATCH}, "value"),
        Input({"type": "blk-fill-color-picker-pt", "index": MATCH}, "value"),
        prevent_initial_call=True,
    )
    def sync_blk_fill_color_pt(text_v, picker_v):
        ctx = dash.callback_context
        trig = (ctx.triggered[0]["prop_id"] if ctx.triggered else "")
        if "picker" in trig:
            v = _normalize_hex_color(picker_v, "#000000")
        else:
            v = _normalize_hex_color(text_v, "#000000")
        return v, v

    @app.callback(
        Output({"type": "blk-type-swatch", "index": MATCH}, "style"),
        Input({"type": "blk-fill-color-text-pt", "index": MATCH}, "value"),
        State({"type": "blk-type-swatch", "index": MATCH}, "style"),
        prevent_initial_call=True,
    )
    def sync_blk_type_swatch(fill_color, cur_style):
        # Keep existing swatch styling, only update its fill color.
        st = dict(cur_style or {})
        st["backgroundColor"] = _normalize_hex_color(fill_color, st.get("backgroundColor", "#999999"))
        return st

    @app.callback(
        Output({"type": "blk-border-color-onoff", "index": MATCH}, "children"),
        Output({"type": "blk-border-color-onoff", "index": MATCH}, "style"),
        Output({"type": "blk-border-color-picker-pt", "index": MATCH}, "disabled"),
        Output({"type": "blk-border-color-text-pt", "index": MATCH}, "disabled"),
        Input({"type": "blk-border-color-onoff", "index": MATCH}, "n_clicks"),
        prevent_initial_call=True,
    )
    def toggle_blk_border_color_onoff(nc):
        on = (nc or 0) % 2 == 1
        label = "On" if on else "Off"
        style = {"minWidth": "48px", "fontSize": "28px", "padding": "2px 6px",
                 "borderRadius": "4px", "border": "2px solid #CCC",
                 "cursor": "pointer"}
        if on:
            style["backgroundColor"] = "#C8E6C9"
            style["color"] = "#2E7D32"
            style["fontWeight"] = "700"
        else:
            style["backgroundColor"] = "#F0F0F0"
            style["color"] = "#333"
            style["fontWeight"] = "400"
        return label, style, not on, not on

    @app.callback(
        Output({"type": "blk-border-color-text-pt", "index": MATCH}, "value"),
        Output({"type": "blk-border-color-picker-pt", "index": MATCH}, "value"),
        Input({"type": "blk-border-color-text-pt", "index": MATCH}, "value"),
        Input({"type": "blk-border-color-picker-pt", "index": MATCH}, "value"),
        prevent_initial_call=True,
    )
    def sync_blk_border_color_pt(text_v, picker_v):
        ctx = dash.callback_context
        trig = (ctx.triggered[0]["prop_id"] if ctx.triggered else "")
        if "picker" in trig:
            v = _normalize_hex_color(picker_v, "#000000")
        else:
            v = _normalize_hex_color(text_v, "#000000")
        return v, v

    @app.callback(
        Output({"type": "blk-text-color-onoff", "index": MATCH}, "children"),
        Output({"type": "blk-text-color-onoff", "index": MATCH}, "style"),
        Output({"type": "blk-text-color-picker-pt", "index": MATCH}, "disabled"),
        Output({"type": "blk-text-color-text-pt", "index": MATCH}, "disabled"),
        Input({"type": "blk-text-color-onoff", "index": MATCH}, "n_clicks"),
        prevent_initial_call=True,
    )
    def toggle_blk_text_color_onoff(nc):
        on = (nc or 0) % 2 == 1
        label = "On" if on else "Off"
        style = {"minWidth": "48px", "fontSize": "28px", "padding": "2px 6px",
                 "borderRadius": "4px", "border": "2px solid #CCC",
                 "cursor": "pointer"}
        if on:
            style["backgroundColor"] = "#C8E6C9"
            style["color"] = "#2E7D32"
            style["fontWeight"] = "700"
        else:
            style["backgroundColor"] = "#F0F0F0"
            style["color"] = "#333"
            style["fontWeight"] = "400"
        return label, style, not on, not on

    @app.callback(
        Output({"type": "blk-text-color-text-pt", "index": MATCH}, "value"),
        Output({"type": "blk-text-color-picker-pt", "index": MATCH}, "value"),
        Input({"type": "blk-text-color-text-pt", "index": MATCH}, "value"),
        Input({"type": "blk-text-color-picker-pt", "index": MATCH}, "value"),
        prevent_initial_call=True,
    )
    def sync_blk_text_color_pt(text_v, picker_v):
        ctx = dash.callback_context
        trig = (ctx.triggered[0]["prop_id"] if ctx.triggered else "")
        if "picker" in trig:
            v = _normalize_hex_color(picker_v, "#000000")
        else:
            v = _normalize_hex_color(text_v, "#000000")
        return v, v

    @app.callback(
        Output({"type": "blk-fill-opacity-onoff", "index": MATCH}, "children"),
        Output({"type": "blk-fill-opacity-onoff", "index": MATCH}, "style"),
        Output({"type": "blk-fill-opacity-slider-pt", "index": MATCH}, "disabled"),
        Output({"type": "blk-fill-opacity-input-pt", "index": MATCH}, "disabled"),
        Input({"type": "blk-fill-opacity-onoff", "index": MATCH}, "n_clicks"),
        prevent_initial_call=True,
    )
    def toggle_blk_fill_opacity_onoff(nc):
        on = (nc or 0) % 2 == 1
        label = "On" if on else "Off"
        style = {"minWidth": "48px", "fontSize": "28px", "padding": "2px 6px",
                 "borderRadius": "4px", "border": "2px solid #CCC",
                 "cursor": "pointer"}
        if on:
            style["backgroundColor"] = "#C8E6C9"
            style["color"] = "#2E7D32"
            style["fontWeight"] = "700"
        else:
            style["backgroundColor"] = "#F0F0F0"
            style["color"] = "#333"
            style["fontWeight"] = "400"
        return label, style, not on, not on

    @app.callback(
        Output({"type": "blk-fill-opacity-input-pt", "index": MATCH}, "value"),
        Output({"type": "blk-fill-opacity-slider-pt", "index": MATCH}, "value"),
        Input({"type": "blk-fill-opacity-input-pt", "index": MATCH}, "value"),
        Input({"type": "blk-fill-opacity-slider-pt", "index": MATCH}, "value"),
        prevent_initial_call=True,
    )
    def sync_blk_fill_opacity_pt(v_in, v_slider):
        ctx = dash.callback_context
        trig = (ctx.triggered[0]["prop_id"] if ctx.triggered else "")
        v = v_in if "input" in trig else v_slider
        if v is None: v = 0
        return v, v

    @app.callback(
        Output({"type": "blk-border-opacity-onoff", "index": MATCH}, "children"),
        Output({"type": "blk-border-opacity-onoff", "index": MATCH}, "style"),
        Output({"type": "blk-border-opacity-slider-pt", "index": MATCH}, "disabled"),
        Output({"type": "blk-border-opacity-input-pt", "index": MATCH}, "disabled"),
        Input({"type": "blk-border-opacity-onoff", "index": MATCH}, "n_clicks"),
        prevent_initial_call=True,
    )
    def toggle_blk_border_opacity_onoff(nc):
        on = (nc or 0) % 2 == 1
        label = "On" if on else "Off"
        style = {"minWidth": "48px", "fontSize": "28px", "padding": "2px 6px",
                 "borderRadius": "4px", "border": "2px solid #CCC",
                 "cursor": "pointer"}
        if on:
            style["backgroundColor"] = "#C8E6C9"
            style["color"] = "#2E7D32"
            style["fontWeight"] = "700"
        else:
            style["backgroundColor"] = "#F0F0F0"
            style["color"] = "#333"
            style["fontWeight"] = "400"
        return label, style, not on, not on

    @app.callback(
        Output({"type": "blk-border-opacity-input-pt", "index": MATCH}, "value"),
        Output({"type": "blk-border-opacity-slider-pt", "index": MATCH}, "value"),
        Input({"type": "blk-border-opacity-input-pt", "index": MATCH}, "value"),
        Input({"type": "blk-border-opacity-slider-pt", "index": MATCH}, "value"),
        prevent_initial_call=True,
    )
    def sync_blk_border_opacity_pt(v_in, v_slider):
        ctx = dash.callback_context
        trig = (ctx.triggered[0]["prop_id"] if ctx.triggered else "")
        v = v_in if "input" in trig else v_slider
        if v is None: v = 0
        return v, v

    @app.callback(
        Output({"type": "blk-border-width-onoff", "index": MATCH}, "children"),
        Output({"type": "blk-border-width-onoff", "index": MATCH}, "style"),
        Output({"type": "blk-border-width-slider-pt", "index": MATCH}, "disabled"),
        Output({"type": "blk-border-width-input-pt", "index": MATCH}, "disabled"),
        Input({"type": "blk-border-width-onoff", "index": MATCH}, "n_clicks"),
        prevent_initial_call=True,
    )
    def toggle_blk_border_width_onoff(nc):
        on = (nc or 0) % 2 == 1
        label = "On" if on else "Off"
        style = {"minWidth": "48px", "fontSize": "28px", "padding": "2px 6px",
                 "borderRadius": "4px", "border": "2px solid #CCC",
                 "cursor": "pointer"}
        if on:
            style["backgroundColor"] = "#C8E6C9"
            style["color"] = "#2E7D32"
            style["fontWeight"] = "700"
        else:
            style["backgroundColor"] = "#F0F0F0"
            style["color"] = "#333"
            style["fontWeight"] = "400"
        return label, style, not on, not on

    @app.callback(
        Output({"type": "blk-border-width-input-pt", "index": MATCH}, "value"),
        Output({"type": "blk-border-width-slider-pt", "index": MATCH}, "value"),
        Input({"type": "blk-border-width-input-pt", "index": MATCH}, "value"),
        Input({"type": "blk-border-width-slider-pt", "index": MATCH}, "value"),
        prevent_initial_call=True,
    )
    def sync_blk_border_width_pt(v_in, v_slider):
        ctx = dash.callback_context
        trig = (ctx.triggered[0]["prop_id"] if ctx.triggered else "")
        v = v_in if "input" in trig else v_slider
        if v is None: v = 0
        return v, v

    @app.callback(
        Output({"type": "blk-text-size-onoff", "index": MATCH}, "children"),
        Output({"type": "blk-text-size-onoff", "index": MATCH}, "style"),
        Output({"type": "blk-text-size-slider-pt", "index": MATCH}, "disabled"),
        Output({"type": "blk-text-size-input-pt", "index": MATCH}, "disabled"),
        Input({"type": "blk-text-size-onoff", "index": MATCH}, "n_clicks"),
        prevent_initial_call=True,
    )
    def toggle_blk_text_size_onoff(nc):
        on = (nc or 0) % 2 == 1
        label = "On" if on else "Off"
        style = {"minWidth": "48px", "fontSize": "28px", "padding": "2px 6px",
                 "borderRadius": "4px", "border": "2px solid #CCC",
                 "cursor": "pointer"}
        if on:
            style["backgroundColor"] = "#C8E6C9"
            style["color"] = "#2E7D32"
            style["fontWeight"] = "700"
        else:
            style["backgroundColor"] = "#F0F0F0"
            style["color"] = "#333"
            style["fontWeight"] = "400"
        return label, style, not on, not on

    @app.callback(
        Output({"type": "blk-text-size-input-pt", "index": MATCH}, "value"),
        Output({"type": "blk-text-size-slider-pt", "index": MATCH}, "value"),
        Input({"type": "blk-text-size-input-pt", "index": MATCH}, "value"),
        Input({"type": "blk-text-size-slider-pt", "index": MATCH}, "value"),
        prevent_initial_call=True,
    )
    def sync_blk_text_size_pt(v_in, v_slider):
        ctx = dash.callback_context
        trig = (ctx.triggered[0]["prop_id"] if ctx.triggered else "")
        v = v_in if "input" in trig else v_slider
        if v is None: v = 0
        return v, v

    @app.callback(
        Output({"type": "blk-line-spacing-onoff", "index": MATCH}, "children"),
        Output({"type": "blk-line-spacing-onoff", "index": MATCH}, "style"),
        Output({"type": "blk-line-spacing-slider-pt", "index": MATCH}, "disabled"),
        Output({"type": "blk-line-spacing-input-pt", "index": MATCH}, "disabled"),
        Input({"type": "blk-line-spacing-onoff", "index": MATCH}, "n_clicks"),
        prevent_initial_call=True,
    )
    def toggle_blk_line_spacing_onoff(nc):
        on = (nc or 0) % 2 == 1
        label = "On" if on else "Off"
        style = {"minWidth": "48px", "fontSize": "28px", "padding": "2px 6px",
                 "borderRadius": "4px", "border": "2px solid #CCC",
                 "cursor": "pointer"}
        if on:
            style["backgroundColor"] = "#C8E6C9"
            style["color"] = "#2E7D32"
            style["fontWeight"] = "700"
        else:
            style["backgroundColor"] = "#F0F0F0"
            style["color"] = "#333"
            style["fontWeight"] = "400"
        return label, style, not on, not on

    @app.callback(
        Output({"type": "blk-line-spacing-input-pt", "index": MATCH}, "value"),
        Output({"type": "blk-line-spacing-slider-pt", "index": MATCH}, "value"),
        Input({"type": "blk-line-spacing-input-pt", "index": MATCH}, "value"),
        Input({"type": "blk-line-spacing-slider-pt", "index": MATCH}, "value"),
        prevent_initial_call=True,
    )
    def sync_blk_line_spacing_pt(v_in, v_slider):
        ctx = dash.callback_context
        trig = (ctx.triggered[0]["prop_id"] if ctx.triggered else "")
        v = v_in if "input" in trig else v_slider
        if v is None: v = 0
        return v, v

    @app.callback(
        Output({"type": "blk-text-rotation-onoff", "index": MATCH}, "children"),
        Output({"type": "blk-text-rotation-onoff", "index": MATCH}, "style"),
        Output({"type": "blk-text-rotation-slider-pt", "index": MATCH}, "disabled"),
        Output({"type": "blk-text-rotation-input-pt", "index": MATCH}, "disabled"),
        Input({"type": "blk-text-rotation-onoff", "index": MATCH}, "n_clicks"),
        prevent_initial_call=True,
    )
    def toggle_blk_text_rotation_onoff(nc):
        on = (nc or 0) % 2 == 1
        label = "On" if on else "Off"
        style = {"minWidth": "48px", "fontSize": "28px", "padding": "2px 6px",
                 "borderRadius": "4px", "border": "2px solid #CCC",
                 "cursor": "pointer"}
        if on:
            style["backgroundColor"] = "#C8E6C9"
            style["color"] = "#2E7D32"
            style["fontWeight"] = "700"
        else:
            style["backgroundColor"] = "#F0F0F0"
            style["color"] = "#333"
            style["fontWeight"] = "400"
        return label, style, not on, not on

    @app.callback(
        Output({"type": "blk-text-rotation-input-pt", "index": MATCH}, "value"),
        Output({"type": "blk-text-rotation-slider-pt", "index": MATCH}, "value"),
        Input({"type": "blk-text-rotation-input-pt", "index": MATCH}, "value"),
        Input({"type": "blk-text-rotation-slider-pt", "index": MATCH}, "value"),
        prevent_initial_call=True,
    )
    def sync_blk_text_rotation_pt(v_in, v_slider):
        ctx = dash.callback_context
        trig = (ctx.triggered[0]["prop_id"] if ctx.triggered else "")
        v = v_in if "input" in trig else v_slider
        if v is None: v = 0
        return v, v

    @app.callback(
        Output({"type": "blk-line-style-onoff", "index": MATCH}, "children"),
        Output({"type": "blk-line-style-onoff", "index": MATCH}, "style"),
        Output({"type": "blk-line-style-pt", "index": MATCH}, "disabled"),
        Input({"type": "blk-line-style-onoff", "index": MATCH}, "n_clicks"),
        prevent_initial_call=True,
    )
    def toggle_blk_line_style_onoff(nc):
        on = (nc or 0) % 2 == 1
        label = "On" if on else "Off"
        style = {"minWidth": "48px", "fontSize": "28px", "padding": "2px 6px",
                 "borderRadius": "4px", "border": "2px solid #CCC",
                 "cursor": "pointer"}
        if on:
            style["backgroundColor"] = "#C8E6C9"
            style["color"] = "#2E7D32"
            style["fontWeight"] = "700"
        else:
            style["backgroundColor"] = "#F0F0F0"
            style["color"] = "#333"
            style["fontWeight"] = "400"
        return label, style, not on

    @app.callback(
        Output({"type": "blk-text-font-onoff", "index": MATCH}, "children"),
        Output({"type": "blk-text-font-onoff", "index": MATCH}, "style"),
        Output({"type": "blk-text-font-pt", "index": MATCH}, "disabled"),
        Input({"type": "blk-text-font-onoff", "index": MATCH}, "n_clicks"),
        prevent_initial_call=True,
    )
    def toggle_blk_text_font_onoff(nc):
        on = (nc or 0) % 2 == 1
        label = "On" if on else "Off"
        style = {"minWidth": "48px", "fontSize": "28px", "padding": "2px 6px",
                 "borderRadius": "4px", "border": "2px solid #CCC",
                 "cursor": "pointer"}
        if on:
            style["backgroundColor"] = "#C8E6C9"
            style["color"] = "#2E7D32"
            style["fontWeight"] = "700"
        else:
            style["backgroundColor"] = "#F0F0F0"
            style["color"] = "#333"
            style["fontWeight"] = "400"
        return label, style, not on

    @app.callback(
        Output({"type": "blk-text-align-onoff", "index": MATCH}, "children"),
        Output({"type": "blk-text-align-onoff", "index": MATCH}, "style"),
        Output({"type": "blk-text-align-pt", "index": MATCH}, "disabled"),
        Input({"type": "blk-text-align-onoff", "index": MATCH}, "n_clicks"),
        prevent_initial_call=True,
    )
    def toggle_blk_text_align_onoff(nc):
        on = (nc or 0) % 2 == 1
        label = "On" if on else "Off"
        style = {"minWidth": "48px", "fontSize": "28px", "padding": "2px 6px",
                 "borderRadius": "4px", "border": "2px solid #CCC",
                 "cursor": "pointer"}
        if on:
            style["backgroundColor"] = "#C8E6C9"
            style["color"] = "#2E7D32"
            style["fontWeight"] = "700"
        else:
            style["backgroundColor"] = "#F0F0F0"
            style["color"] = "#333"
            style["fontWeight"] = "400"
        return label, style, not on

    @app.callback(
        Output({"type": "blk-radius-onoff", "index": MATCH}, "children"),
        Output({"type": "blk-radius-onoff", "index": MATCH}, "style"),
        Output({"type": "blk-radius-tl-pt", "index": MATCH}, "disabled"),
        Output({"type": "blk-radius-tr-pt", "index": MATCH}, "disabled"),
        Output({"type": "blk-radius-bl-pt", "index": MATCH}, "disabled"),
        Output({"type": "blk-radius-br-pt", "index": MATCH}, "disabled"),
        Input({"type": "blk-radius-onoff", "index": MATCH}, "n_clicks"),
        prevent_initial_call=True,
    )
    def toggle_blk_radius_onoff(nc):
        on = (nc or 0) % 2 == 1
        label = "On" if on else "Off"
        style = {"minWidth": "48px", "fontSize": "28px", "padding": "2px 6px",
                 "borderRadius": "4px", "border": "2px solid #CCC",
                 "cursor": "pointer"}
        if on:
            style["backgroundColor"] = "#C8E6C9"
            style["color"] = "#2E7D32"
            style["fontWeight"] = "700"
        else:
            style["backgroundColor"] = "#F0F0F0"
            style["color"] = "#333"
            style["fontWeight"] = "400"
        return label, style, not on, not on, not on, not on

    # Master update_style_store_block: reads ALL per-type controls
    # Birth subtypes (EnBirth, ExBirth, HyBirth) inherit from Birth when their On/Off is Off.
    @app.callback(
        Output("store-style", "data", allow_duplicate=True),
        Input({"type": "blk-fill-color-text-pt", "index": ALL}, "value"),
        Input({"type": "blk-fill-color-picker-pt", "index": ALL}, "value"),
        Input({"type": "blk-fill-opacity-slider-pt", "index": ALL}, "value"),
        Input({"type": "blk-border-width-slider-pt", "index": ALL}, "value"),
        Input({"type": "blk-line-style-pt", "index": ALL}, "value"),
        Input({"type": "blk-border-color-text-pt", "index": ALL}, "value"),
        Input({"type": "blk-border-color-picker-pt", "index": ALL}, "value"),
        Input({"type": "blk-border-opacity-slider-pt", "index": ALL}, "value"),
        Input({"type": "blk-radius-tl-pt", "index": ALL}, "value"),
        Input({"type": "blk-radius-tr-pt", "index": ALL}, "value"),
        Input({"type": "blk-radius-bl-pt", "index": ALL}, "value"),
        Input({"type": "blk-radius-br-pt", "index": ALL}, "value"),
        Input({"type": "blk-text-font-pt", "index": ALL}, "value"),
        Input({"type": "blk-text-size-slider-pt", "index": ALL}, "value"),
        Input({"type": "blk-text-color-text-pt", "index": ALL}, "value"),
        Input({"type": "blk-text-color-picker-pt", "index": ALL}, "value"),
        Input({"type": "blk-text-align-pt", "index": ALL}, "value"),
        Input({"type": "blk-line-spacing-slider-pt", "index": ALL}, "value"),
        Input({"type": "blk-text-rotation-slider-pt", "index": ALL}, "value"),
        Input("block-width-slider", "value"),
        Input("block-height-slider", "value"),
        Input("blk-text-pad-l", "value"),
        Input("blk-text-pad-r", "value"),
        Input("blk-text-pad-t", "value"),
        Input("blk-text-pad-b", "value"),
        State("store-style", "data"),
        State("store-block-mode", "data"),
        # On/Off n_clicks for Birth subtype inheritance (13 properties)
        State({"type": "blk-fill-color-onoff", "index": ALL}, "n_clicks"),
        State({"type": "blk-fill-opacity-onoff", "index": ALL}, "n_clicks"),
        State({"type": "blk-border-width-onoff", "index": ALL}, "n_clicks"),
        State({"type": "blk-line-style-onoff", "index": ALL}, "n_clicks"),
        State({"type": "blk-border-color-onoff", "index": ALL}, "n_clicks"),
        State({"type": "blk-border-opacity-onoff", "index": ALL}, "n_clicks"),
        State({"type": "blk-radius-onoff", "index": ALL}, "n_clicks"),
        State({"type": "blk-text-font-onoff", "index": ALL}, "n_clicks"),
        State({"type": "blk-text-size-onoff", "index": ALL}, "n_clicks"),
        State({"type": "blk-text-color-onoff", "index": ALL}, "n_clicks"),
        State({"type": "blk-text-align-onoff", "index": ALL}, "n_clicks"),
        State({"type": "blk-line-spacing-onoff", "index": ALL}, "n_clicks"),
        State({"type": "blk-text-rotation-onoff", "index": ALL}, "n_clicks"),
        prevent_initial_call=True,
    )
    def update_style_store_block(
            fill_color_texts, fill_color_pickers, fill_opacity_sliders,
            border_width_sliders, line_style_vals,
            border_color_texts, border_color_pickers, border_opacity_sliders,
            radius_tls, radius_trs, radius_bls, radius_brs,
            text_font_vals, text_size_sliders,
            text_color_texts, text_color_pickers, text_align_vals,
            line_spacing_sliders, text_rotation_sliders,
            width_v, height_v,
            pad_l_v, pad_r_v, pad_t_v, pad_b_v,
            style_state, block_mode_state,
            oo_fill_color, oo_fill_opacity, oo_border_width, oo_line_style,
            oo_border_color, oo_border_opacity, oo_radius,
            oo_text_font, oo_text_size, oo_text_color, oo_text_align,
            oo_line_spacing, oo_text_rotation,
    ):
        style_state = style_state or {}
        block_state = dict(style_state.get("block", {}) or {})

        widths = dict(block_state.get("widths") or {})
        heights = dict(block_state.get("heights") or {})
        fill_colors = dict(block_state.get("fill_colors") or {})
        fill_opacities = dict(block_state.get("fill_opacities") or {})
        border_colors = dict(block_state.get("border_colors") or {})
        border_opacities = dict(block_state.get("border_opacities") or {})
        border_radii = dict(block_state.get("border_radii") or {})
        line_styles = dict(block_state.get("line_styles") or {})
        text_fonts = dict(block_state.get("text_fonts") or {})
        text_sizes = dict(block_state.get("text_sizes") or {})
        text_colors = dict(block_state.get("text_colors") or {})
        text_aligns = dict(block_state.get("text_aligns") or {})
        line_spacings = dict(block_state.get("line_spacings") or {})
        text_rotations = dict(block_state.get("text_rotations") or {})
        border_widths = dict(block_state.get("border_widths") or {})

        for k in cfg.block_colors.keys():
            widths.setdefault(k, float(cfg.block_width))
            heights.setdefault(k, float(cfg.block_height))
            fill_colors.setdefault(k, cfg.block_colors[k])
            fill_opacities.setdefault(k, 1.0)
            border_colors.setdefault(k, "#222222")
            border_opacities.setdefault(k, 1.0)
            border_radii.setdefault(k, [0, 0, 0, 0])
            line_styles.setdefault(k, "solid")
            text_fonts.setdefault(k, "Arial")
            text_sizes.setdefault(k, int(cfg.block_text_size))
            text_colors.setdefault(k, cfg.block_text_colors.get(k, "#111111"))
            text_aligns.setdefault(k, "center")
            line_spacings.setdefault(k, 0)
            text_rotations.setdefault(k, 90)
            border_widths.setdefault(k, float(cfg.block_border_width))

        # Per-type values with 3-level "last-touch wins" semantics:
        # - Level 1: All
        # - Level 2: Type (e.g., Birth)
        # - Level 3: Birth subtypes (EnBirth/ExBirth/HyBirth)
        #
        # Initial state: every real type uses its own value (implicit touch=1),
        # while parents ("Birth", "All") have implicit touch=0.
        # After interaction: whichever control is adjusted most recently becomes the effective source.
        ctx = dash.callback_context
        trig = (ctx.triggered[0]["prop_id"] if (ctx.triggered and len(ctx.triggered) > 0) else "")

        # Metadata store for touch clocks
        meta = dict(block_state.get("_meta") or {})
        clock = int(meta.get("clock") or 0)
        touch = dict(meta.get("touch") or {})  # prop_key -> {bt -> clock}

        # Map triggered prop_id to (bt, prop_key)
        def _parse_trigger(trig_prop_id):
            # trig_prop_id example: '{"type":"blk-fill-color-text-pt","index":"Birth"}.value'
            if not trig_prop_id:
                return None, None
            try:
                comp, _ = trig_prop_id.rsplit(".", 1)
                tid = json.loads(comp)
                bt = tid.get("index")
                t = tid.get("type", "")
            except Exception:
                return None, None

            # Normalize prop_key per control "type"
            if t in ("blk-fill-color-text-pt", "blk-fill-color-picker-pt"):
                return bt, "fill_colors"
            if t == "blk-fill-opacity-slider-pt":
                return bt, "fill_opacities"
            if t == "blk-border-width-slider-pt":
                return bt, "border_widths"
            if t == "blk-line-style-pt":
                return bt, "line_styles"
            if t in ("blk-border-color-text-pt", "blk-border-color-picker-pt"):
                return bt, "border_colors"
            if t == "blk-border-opacity-slider-pt":
                return bt, "border_opacities"
            if t in ("blk-radius-tl-pt", "blk-radius-tr-pt", "blk-radius-bl-pt", "blk-radius-br-pt"):
                return bt, "border_radii"
            if t == "blk-text-font-pt":
                return bt, "text_fonts"
            if t == "blk-text-size-slider-pt":
                return bt, "text_sizes"
            if t in ("blk-text-color-text-pt", "blk-text-color-picker-pt"):
                return bt, "text_colors"
            if t == "blk-text-align-pt":
                return bt, "text_aligns"
            if t == "blk-line-spacing-slider-pt":
                return bt, "line_spacings"
            if t == "blk-text-rotation-slider-pt":
                return bt, "text_rotations"
            return bt, None

        bt_trig, prop_trig = _parse_trigger(trig)
        if bt_trig is not None and prop_trig is not None:
            clock += 1
            tmap = dict(touch.get(prop_trig) or {})
            tmap[str(bt_trig)] = int(clock)
            touch[prop_trig] = tmap

        # Helpers for implicit/default touch clocks
        def _implicit_touch(prop_key, bt_name):
            # Real types start at 1 (self wins initially). Parents start at 0.
            if bt_name == "All":
                return 0
            if prop_key in ("fill_colors", "fill_opacities", "border_widths", "line_styles",
                            "border_colors", "border_opacities", "border_radii",
                            "text_fonts", "text_sizes", "text_colors", "text_aligns",
                            "line_spacings", "text_rotations"):
                # Birth is a real type (it has its own block); subtypes are also real types.
                # However, for the hierarchical override, Birth acts as a parent only for its subtypes.
                return 1
            return 1

        def _touch_time(prop_key, bt_name):
            tmap = dict(touch.get(prop_key) or {})
            return int(tmap.get(str(bt_name)) or _implicit_touch(prop_key, bt_name))

        def _pick_source(prop_key, bt_name):
            # Candidate chain: self -> Birth (if subtype) -> All
            candidates = [bt_name]
            if bt_name in ("EnBirth", "ExBirth", "HyBirth"):
                candidates.append("Birth")
            candidates.append("All")
            best = candidates[0]
            best_t = _touch_time(prop_key, candidates[0])
            for c in candidates[1:]:
                tt = _touch_time(prop_key, c)
                if tt >= best_t:
                    best = c
                    best_t = tt
            return best

        # Read ALL control values into raw maps (including "All" and "Birth")
        raw = {
            "fill_colors": {},
            "fill_opacities": {},
            "border_widths": {},
            "line_styles": {},
            "border_colors": {},
            "border_opacities": {},
            "border_radii": {},
            "text_fonts": {},
            "text_sizes": {},
            "text_colors": {},
            "text_aligns": {},
            "line_spacings": {},
            "text_rotations": {},
        }

        for i, bt in enumerate(BLOCK_TYPE_UI_ORDER):
            # Fill color
            if i < len(fill_color_texts):
                raw["fill_colors"][bt] = _normalize_hex_color(fill_color_texts[i], cfg.block_colors.get(bt, "#CCCCCC"))
            # Fill opacity
            if i < len(fill_opacity_sliders):
                raw["fill_opacities"][bt] = _clamp01(fill_opacity_sliders[i], 1.0)
            # Border width
            if i < len(border_width_sliders):
                raw["border_widths"][bt] = _clamp_border_width(border_width_sliders[i], cfg.block_border_width)
            # Line style
            if i < len(line_style_vals):
                raw["line_styles"][bt] = line_style_vals[i] or "solid"
            # Border color
            if i < len(border_color_texts):
                raw["border_colors"][bt] = _normalize_hex_color(border_color_texts[i], "#222222")
            # Border opacity
            if i < len(border_opacity_sliders):
                raw["border_opacities"][bt] = _clamp01(border_opacity_sliders[i], 1.0)
            # Radii
            if i < len(radius_tls):
                raw["border_radii"][bt] = [_clamp_radius(radius_tls[i]), _clamp_radius(radius_trs[i]),
                                           _clamp_radius(radius_brs[i]), _clamp_radius(radius_bls[i])]
            # Text font
            if i < len(text_font_vals):
                raw["text_fonts"][bt] = text_font_vals[i] or "Arial"
            # Text size
            if i < len(text_size_sliders):
                raw["text_sizes"][bt] = _clamp_text_size(text_size_sliders[i], cfg.block_text_size)
            # Text color
            if i < len(text_color_texts):
                raw["text_colors"][bt] = _normalize_hex_color(text_color_texts[i],
                                                              cfg.block_text_colors.get(bt, "#111111"))
            # Text align
            if i < len(text_align_vals):
                raw["text_aligns"][bt] = text_align_vals[i] or "center"
            # Line spacing
            if i < len(line_spacing_sliders):
                try:
                    ls = float(line_spacing_sliders[i])
                    ls = max(-20.0, min(200.0, ls))
                except Exception:
                    ls = 0.0
                raw["line_spacings"][bt] = ls
            # Text rotation
            if i < len(text_rotation_sliders):
                try:
                    rot = int(text_rotation_sliders[i])
                    rot = max(-180, min(180, rot))
                except Exception:
                    rot = 0
                raw["text_rotations"][bt] = rot

        # Apply effective values to real block types (exclude "All" pseudo-type)
        for bt in cfg.block_colors.keys():
            # For each property, select the most recently touched source among {self, Birth, All}
            src = _pick_source("fill_colors", bt)
            if src in raw["fill_colors"]:
                fill_colors[bt] = raw["fill_colors"][src]

            src = _pick_source("fill_opacities", bt)
            if src in raw["fill_opacities"]:
                fill_opacities[bt] = raw["fill_opacities"][src]

            src = _pick_source("border_widths", bt)
            if src in raw["border_widths"]:
                border_widths[bt] = raw["border_widths"][src]

            src = _pick_source("line_styles", bt)
            if src in raw["line_styles"]:
                line_styles[bt] = raw["line_styles"][src]

            src = _pick_source("border_colors", bt)
            if src in raw["border_colors"]:
                border_colors[bt] = raw["border_colors"][src]

            src = _pick_source("border_opacities", bt)
            if src in raw["border_opacities"]:
                border_opacities[bt] = raw["border_opacities"][src]

            src = _pick_source("border_radii", bt)
            if src in raw["border_radii"]:
                border_radii[bt] = raw["border_radii"][src]

            src = _pick_source("text_fonts", bt)
            if src in raw["text_fonts"]:
                text_fonts[bt] = raw["text_fonts"][src]

            src = _pick_source("text_sizes", bt)
            if src in raw["text_sizes"]:
                text_sizes[bt] = raw["text_sizes"][src]

            src = _pick_source("text_colors", bt)
            if src in raw["text_colors"]:
                text_colors[bt] = raw["text_colors"][src]

            src = _pick_source("text_aligns", bt)
            if src in raw["text_aligns"]:
                text_aligns[bt] = raw["text_aligns"][src]

            src = _pick_source("line_spacings", bt)
            if src in raw["line_spacings"]:
                line_spacings[bt] = raw["line_spacings"][src]

            src = _pick_source("text_rotations", bt)
            if src in raw["text_rotations"]:
                text_rotations[bt] = raw["text_rotations"][src]

        # Persist metadata
        meta["clock"] = int(clock)
        meta["touch"] = touch
        block_state["_meta"] = meta
        # Global width/height
        current_block_mode = (block_mode_state or {}).get("mode", "existence")
        if current_block_mode == "existence":
            w = _clamp_block_width(width_v, cfg.block_width)
            for k in widths:
                widths[k] = w
        h = _clamp_block_height(height_v, cfg.block_height)
        for k in heights:
            heights[k] = h

        block_state["widths"] = widths
        block_state["heights"] = heights
        block_state["fill_colors"] = fill_colors
        block_state["fill_opacities"] = fill_opacities
        block_state["border_colors"] = border_colors
        block_state["border_opacities"] = border_opacities
        block_state["border_radii"] = border_radii
        block_state["line_styles"] = line_styles
        block_state["text_fonts"] = text_fonts
        block_state["text_sizes"] = text_sizes
        block_state["text_colors"] = text_colors
        block_state["text_aligns"] = text_aligns
        block_state["line_spacings"] = line_spacings
        block_state["text_rotations"] = text_rotations
        block_state["border_widths"] = border_widths

        # Text padding (L, R, T, B) — global from ALL panel
        def _clamp_pad(v, fallback=0.0):
            try:
                return float(max(0.0, min(5.0, float(v))))
            except Exception:
                return float(fallback)

        block_state["text_pads"] = {
            "l": _clamp_pad(pad_l_v),
            "r": _clamp_pad(pad_r_v),
            "t": _clamp_pad(pad_t_v),
            "b": _clamp_pad(pad_b_v),
        }

        style_state["block"] = block_state

        return style_state

    # Style panel (Group) callbacks

    # Sync group line width (input <-> slider)
    @app.callback(
        Output("group-line-width-input", "value"),
        Output("group-line-width-slider", "value"),
        Input("group-line-width-input", "value"),
        Input("group-line-width-slider", "value"),
        prevent_initial_call=True,
    )
    def sync_group_line_width(v_in, v_slider):
        ctx = dash.callback_context
        trig = (ctx.triggered[0]["prop_id"] if ctx.triggered else "")

        def _clamp_lw(x):
            try:
                v = float(x)
            except Exception:
                v = float(cfg.group_enclosure_line_width)
            return float(max(0.5, min(10.0, v)))

        if trig.startswith("group-line-width-input"):
            v = _clamp_lw(v_in)
        else:
            v = _clamp_lw(v_slider)
        return v, v

    # Sync group opacity (input <-> slider)
    @app.callback(
        Output("group-opacity-input", "value"),
        Output("group-opacity-slider", "value"),
        Input("group-opacity-input", "value"),
        Input("group-opacity-slider", "value"),
        prevent_initial_call=True,
    )
    def sync_group_opacity(v_in, v_slider):
        ctx = dash.callback_context
        trig = (ctx.triggered[0]["prop_id"] if ctx.triggered else "")

        if trig.startswith("group-opacity-input"):
            v = _clamp01(v_in, cfg.group_enclosure_opacity)
        else:
            v = _clamp01(v_slider, cfg.group_enclosure_opacity)
        return v, v

    # Sync group color (text <-> picker)
    @app.callback(
        Output("group-color-text", "value"),
        Output("group-color-picker", "value"),
        Input("group-color-text", "value"),
        Input("group-color-picker", "value"),
        prevent_initial_call=True,
    )
    def sync_group_color(text_v, picker_v):
        ctx = dash.callback_context
        trig = (ctx.triggered[0]["prop_id"] if ctx.triggered else "")

        if trig.startswith("group-color-picker"):
            v = _normalize_hex_color(picker_v, "#000000")
        else:
            v = _normalize_hex_color(text_v, "#000000")
        return v, v

    # Sync group label size (input <-> slider)
    @app.callback(
        Output("group-label-size-input", "value"),
        Output("group-label-size-slider", "value"),
        Input("group-label-size-input", "value"),
        Input("group-label-size-slider", "value"),
        prevent_initial_call=True,
    )
    def sync_group_label_size(v_in, v_slider):
        ctx = dash.callback_context
        trig = (ctx.triggered[0]["prop_id"] if ctx.triggered else "")

        def _clamp_sz(x):
            try:
                v = int(x)
            except Exception:
                v = 12
            return int(max(6, min(48, v)))

        if trig.startswith("group-label-size-input"):
            v = _clamp_sz(v_in)
        else:
            v = _clamp_sz(v_slider)
        return v, v

    # Sync group label color (text <-> picker)
    @app.callback(
        Output("group-label-color-text", "value"),
        Output("group-label-color-picker", "value"),
        Input("group-label-color-text", "value"),
        Input("group-label-color-picker", "value"),
        prevent_initial_call=True,
    )
    def sync_group_label_color(text_v, picker_v):
        ctx = dash.callback_context
        trig = (ctx.triggered[0]["prop_id"] if ctx.triggered else "")

        if trig.startswith("group-label-color-picker"):
            v = _normalize_hex_color(picker_v, "#000000")
        else:
            v = _normalize_hex_color(text_v, "#000000")
        return v, v

    # Sync group label distance (input <-> slider)
    @app.callback(
        Output("group-label-offset-x-input", "value"),
        Output("group-label-offset-x-slider", "value"),
        Input("group-label-offset-x-input", "value"),
        Input("group-label-offset-x-slider", "value"),
        prevent_initial_call=True,
    )
    def sync_group_label_offset_x(v_in, v_slider):
        ctx = dash.callback_context
        trig = (ctx.triggered[0]["prop_id"] if ctx.triggered else "")

        def _clamp_dx(x):
            try:
                v = float(x)
            except Exception:
                v = float(cfg.enclosure_group_label_offset_x)
            return float(max(0.0, min(2.0, v)))

        if trig.startswith("group-label-offset-x-input"):
            v = _clamp_dx(v_in)
        else:
            v = _clamp_dx(v_slider)
        return v, v

    # Sync group inner gap (input <-> slider)
    @app.callback(
        Output("group-inner-gap-input", "value"),
        Output("group-inner-gap-slider", "value"),
        Input("group-inner-gap-input", "value"),
        Input("group-inner-gap-slider", "value"),
        prevent_initial_call=True,
    )
    def sync_group_inner_gap(v_in, v_slider):
        ctx = dash.callback_context
        trig = (ctx.triggered[0]["prop_id"] if ctx.triggered else "")

        def _clamp_inner(x):
            try:
                v = float(x)
            except Exception:
                v = float(cfg.inner_gap)
            return float(max(0.0, min(1.0, v)))

        if trig.startswith("group-inner-gap-input"):
            v = _clamp_inner(v_in)
        else:
            v = _clamp_inner(v_slider)
        return v, v

    # Sync group outer gap (input <-> slider)
    @app.callback(
        Output("group-outer-gap-input", "value"),
        Output("group-outer-gap-slider", "value"),
        Input("group-outer-gap-input", "value"),
        Input("group-outer-gap-slider", "value"),
        prevent_initial_call=True,
    )
    def sync_group_outer_gap(v_in, v_slider):
        ctx = dash.callback_context
        trig = (ctx.triggered[0]["prop_id"] if ctx.triggered else "")

        def _clamp_outer(x):
            try:
                v = float(x)
            except Exception:
                v = float(cfg.outer_gap)
            return float(max(0.0, min(10.0, v)))

        if trig.startswith("group-outer-gap-input"):
            v = _clamp_outer(v_in)
        else:
            v = _clamp_outer(v_slider)
        return v, v

        # Sync group name rotation slider and input

    @app.callback(
        Output("group-text-rotation", "value"),
        Output("group-text-rotation-slider", "value"),
        Input("group-text-rotation", "value"),
        Input("group-text-rotation-slider", "value"),
        prevent_initial_call=True,
    )
    def sync_group_text_rotation(v_in, v_slider):
        ctx = dash.callback_context
        trig = (ctx.triggered[0]["prop_id"] if ctx.triggered else "")
        v = v_in if trig.startswith("group-text-rotation") and not trig.startswith(
            "group-text-rotation-slider") else v_slider
        if v is None:
            v = 90
        try:
            v = int(v)
        except Exception:
            v = 90
        v = max(-180, min(180, v))
        return v, v

    # Sync group label line spacing slider and input
    @app.callback(
        Output("group-label-line-spacing-input", "value"),
        Output("group-label-line-spacing-slider", "value"),
        Input("group-label-line-spacing-input", "value"),
        Input("group-label-line-spacing-slider", "value"),
        prevent_initial_call=True,
    )
    def sync_group_label_line_spacing(v_in, v_slider):
        ctx = dash.callback_context
        trig = (ctx.triggered[0]["prop_id"] if ctx.triggered else "")
        v = v_in if trig.startswith("group-label-line-spacing-input") else v_slider
        if v is None:
            v = 0
        try:
            v = float(v)
        except Exception:
            v = 0
        v = max(-20, min(200, v))
        return v, v

    # Update store-style with group settings
    @app.callback(
        Output("store-style", "data", allow_duplicate=True),
        Input("group-line-width-slider", "value"),
        Input("group-opacity-slider", "value"),
        Input("group-color-text", "value"),
        Input("group-color-picker", "value"),
        Input("group-radius-tl", "value"),
        Input("group-radius-tr", "value"),
        Input("group-radius-bl", "value"),
        Input("group-radius-br", "value"),
        Input("group-line-style", "value"),
        Input("group-label-font", "value"),
        Input("group-label-size-slider", "value"),
        Input("group-label-color-text", "value"),
        Input("group-label-color-picker", "value"),
        Input("group-label-offset-x-slider", "value"),
        Input("group-text-rotation", "value"),
        Input("group-inner-gap-slider", "value"),
        Input("group-outer-gap-slider", "value"),
        Input("group-label-line-spacing-slider", "value"),
        Input("group-label-pad-b", "value"),
        Input("group-label-pad-t", "value"),
        State("store-style", "data"),
        prevent_initial_call=True,
    )
    def update_style_store_group(
            line_width_v, opacity_v, color_text, color_picker,
            radius_tl, radius_tr, radius_bl, radius_br, line_style_v,
            label_font_v, label_size_v, label_color_text, label_color_picker, label_offset_x_v,
            label_rotation_v,
            inner_gap_v, outer_gap_v,
            label_line_spacing_v, label_pad_b_v, label_pad_t_v,
            style_state
    ):
        style_state = style_state or {}
        group_state = dict(style_state.get("group", {}) or {})

        ctx = dash.callback_context
        trig = (ctx.triggered[0]["prop_id"] if ctx.triggered else "")

        if "group-line-width" in trig:
            try:
                group_state["line_width"] = float(max(0.5, min(10.0, float(line_width_v))))
            except Exception:
                group_state["line_width"] = float(cfg.group_enclosure_line_width)

        if "group-opacity" in trig:
            group_state["opacity"] = _clamp01(opacity_v, cfg.group_enclosure_opacity)

        if "group-color" in trig:
            v = _normalize_hex_color(color_picker if "picker" in trig else color_text, "#000000")
            group_state["color"] = v

        if "group-radius" in trig:
            r = [_clamp_radius(radius_tl), _clamp_radius(radius_tr),
                 _clamp_radius(radius_br), _clamp_radius(radius_bl)]
            group_state["radii"] = r

        if "group-line-style" in trig:
            group_state["line_style"] = line_style_v or "solid"

        if "group-label-font" in trig:
            group_state["label_font"] = label_font_v or "Arial"

        if "group-label-size" in trig:
            try:
                group_state["label_size"] = int(max(6, min(48, int(label_size_v))))
            except Exception:
                group_state["label_size"] = 8

        if "group-label-color" in trig:
            v = _normalize_hex_color(label_color_picker if "picker" in trig else label_color_text, "#000000")
            group_state["label_color"] = v

        if "group-label-offset-x" in trig:
            try:
                group_state["label_offset_x"] = float(max(0.0, min(2.0, float(label_offset_x_v))))
            except Exception:
                group_state["label_offset_x"] = float(cfg.enclosure_group_label_offset_x)

        if "group-text-rotation" in trig:
            try:
                rot = int(label_rotation_v)
                rot = max(-180, min(180, rot))
            except Exception:
                rot = 0
            group_state["label_rotation"] = rot

        if "group-label-line-spacing" in trig:
            try:
                ls = float(max(-20, min(200, float(label_line_spacing_v))))
            except Exception:
                ls = 0
            group_state["label_line_spacing"] = ls

        if "group-label-pad-b" in trig:
            try:
                group_state["label_pad_b"] = float(max(0.0, min(5.0, float(label_pad_b_v))))
            except Exception:
                group_state["label_pad_b"] = 0.0

        if "group-label-pad-t" in trig:
            try:
                group_state["label_pad_t"] = float(max(0.0, min(5.0, float(label_pad_t_v))))
            except Exception:
                group_state["label_pad_t"] = 0.0

        if "group-inner-gap" in trig:
            layout_state = dict(style_state.get("layout", {}) or {})
            try:
                layout_state["inner_gap"] = float(max(0.0, min(1.0, float(inner_gap_v))))
            except Exception:
                layout_state["inner_gap"] = float(cfg.inner_gap)
            style_state["layout"] = layout_state

        if "group-outer-gap" in trig:
            layout_state = dict(style_state.get("layout", {}) or {})
            try:
                layout_state["outer_gap"] = float(max(0.0, min(10.0, float(outer_gap_v))))
            except Exception:
                layout_state["outer_gap"] = float(cfg.outer_gap)
            style_state["layout"] = layout_state

        style_state["group"] = group_state
        return style_state

    # Style panel (Supergroup) callbacks - Dynamic

    def _get_default_supergroup_style(level: int) -> Dict[str, Any]:
        """Get default style for a supergroup level."""
        return {
            "pad_left": (0.5 if level == 1 else 0.2),
            "pad_right": (0.5 if level == 1 else 0.2),
            "pad_top": (0.5 if level == 1 else 0.5),
            "pad_bottom": (0.5 if level == 1 else 0.5),
            "enclosure_gap": (2.0 if level == 1 else 3.0),
            "fill_color": "#FFFFFF",
            "fill_opacity": 0.0,
            "border_color": "#000000",
            "border_opacity": (0.5 if level == 1 else 0.4),
            "border_width": (1.0 if level == 1 else 0.7),
            "radii": [0, 0, 0, 0],
            "line_style": "solid",
            "label_font": "Arial",
            "label_size": 8,
            "label_color": "#000000",
            "label_distance": -0.05,
            "label_rotation": 0,
        }

    def _render_supergroup_panel(level: int, style: Dict[str, Any], ctrl_label_style: dict, ctrl_box_style: dict,
                                 is_shown: bool = False, is_enclosure_visible: bool = True,
                                 is_highest_level: bool = False) -> html.Div:
        """Render a single enclosure control panel."""
        s = style or _get_default_supergroup_style(level)
        level_str = str(level)
        btn_text = f"Hide enclosure{level} name" if is_shown else f"Show enclosure{level} name"
        enclosure_btn_text = f"Hide enclosure{level}" if is_enclosure_visible else f"Show enclosure{level}"

        def _safe_float(v, default, min_v=None, max_v=None):
            try:
                if v is None or v == "" or v == "None":
                    v2 = float(default)
                else:
                    v2 = float(v)
            except Exception:
                v2 = float(default)
            if min_v is not None:
                v2 = max(float(min_v), v2)
            if max_v is not None:
                v2 = min(float(max_v), v2)
            return float(v2)

        def _safe_int(v, default, min_v=None, max_v=None):
            try:
                if v is None or v == "" or v == "None":
                    v2 = int(default)
                else:
                    v2 = int(float(v))
            except Exception:
                v2 = int(default)
            if min_v is not None:
                v2 = max(int(min_v), v2)
            if max_v is not None:
                v2 = min(int(max_v), v2)
            return int(v2)

        # Safely get radii values
        radii_raw = s.get("radii")
        radii = radii_raw if radii_raw and isinstance(radii_raw, (list, tuple)) and len(radii_raw) >= 4 else [0, 0, 0,
                                                                                                              0]
        r_tl = int(radii[0] or 0) if radii[0] is not None else 0
        r_tr = int(radii[1] or 0) if radii[1] is not None else 0
        r_br = int(radii[2] or 0) if radii[2] is not None else 0
        r_bl = int(radii[3] or 0) if radii[3] is not None else 0

        pad_left_v = _safe_float(s.get("pad_left"), 0.0, 0.0, 5.0)
        pad_right_v = _safe_float(s.get("pad_right"), 0.0, 0.0, 5.0)
        pad_top_v = _safe_float(s.get("pad_top"), 0.0, 0.0, 5.0)
        pad_bottom_v = _safe_float(s.get("pad_bottom"), 0.0, 0.0, 5.0)
        border_width_v = _safe_float(s.get("border_width"), 1.0, 0.0, 5.0)
        border_opacity_v = _safe_float(s.get("border_opacity"),
                                       _get_default_supergroup_style(level).get("border_opacity", 0.35), 0.0, 1.0)
        label_size_v = _safe_int(s.get("label_size"), 23, 6, 48)
        label_distance_v = _safe_float(s.get("label_distance"), -0.05, -2.0, 2.0)
        label_rotation_v = _safe_int(s.get("label_rotation"), 0, -180, 180)

        border_color_v = str(s.get("border_color") or "#000000")
        label_color_v = str(s.get("label_color") or "#000000")
        line_style_v = str(s.get("line_style") or "solid")
        label_font_v = str(s.get("label_font") or "Arial")

        enclosure_gap_v = _safe_float(s.get("enclosure_gap"),
                                      _get_default_supergroup_style(level).get("enclosure_gap", 2.0), 0.0, 10.0)

        sg_pad_input_style = {**pair_input_style}

        return html.Div([
            html.Label(f"Enclosure{level}", style={**ctrl_label_style, "fontWeight": "700", "marginBottom": "10px"}),
            html.Div(
                style={"display": "flex", "flexDirection": "column", "gap": "10px"},
                children=[
                    # Gap control (distance between sibling enclosures) — only on highest level
                    html.Div(
                        style={"display": "flex" if is_highest_level else "none", "flexDirection": "row",
                               "alignItems": "center", "justifyContent": "flex-start", "gap": "10px"},
                        children=[
                            html.Label("Gap",
                                       style={**ctrl_label_style, "marginBottom": "0px", "width": right_label_w}),
                            html.Div(
                                style=right_ctrl_col_style,
                                children=[
                                    html.Div(
                                        style={"flex": "1", "minWidth": "0"},
                                        children=[
                                            dcc.Slider(id={"type": "sg-enclosure-gap-slider", "level": level_str},
                                                       min=0.0, max=4.0, step=0.05, value=enclosure_gap_v,
                                                       updatemode="mouseup", tooltip={"always_visible": False}),
                                        ],
                                    ),
                                    dcc.Input(debounce=True, id={"type": "sg-enclosure-gap-input", "level": level_str},
                                              type="number", min=0.0, max=10.0, step="any", value=enclosure_gap_v,
                                              style={"width": "100px", **ctrl_box_style}),
                                ],
                            ),
                        ],
                    ),
                    # Hide/Show enclosure button (border visibility)
                    html.Div(
                        style={"boxSizing": "border-box"},
                        children=[
                            html.Button(enclosure_btn_text, id={"type": "btn-sg-visibility-toggle", "level": level_str},
                                        n_clicks=0, style={"width": "100%"}),
                        ],
                    ),
                    html.Div(
                        style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                               "justifyContent": "flex-start", "gap": "10px", "width": "100%"},
                        children=[
                            html.Label("Radius",
                                       style={**ctrl_label_style, "marginBottom": "0px", "width": right_label_w}),
                            html.Div(
                                style=right_ctrl_single_style,
                                children=[
                                    html.Div(
                                        style={"display": "flex", "flexDirection": "column", "gap": "8px"},
                                        children=[
                                            html.Div(
                                                style=pair_two_input_row_style,
                                                children=[
                                                    html.Span("TL", style=pair_small_label_style),
                                                    html.Div(),
                                                    dcc.Input(id={"type": "sg-radius-tl", "level": level_str},
                                                              type="number", min=0, max=50, step="any", value=r_tl,
                                                              debounce=False, style=pair_input_style),
                                                    html.Div(),
                                                    html.Span("TR", style=pair_small_label_style),
                                                    html.Div(),
                                                    dcc.Input(id={"type": "sg-radius-tr", "level": level_str},
                                                              type="number", min=0, max=50, step="any", value=r_tr,
                                                              debounce=False, style=pair_input_style),
                                                ],
                                            ),
                                            html.Div(
                                                style=pair_two_input_row_style,
                                                children=[
                                                    html.Span("BL", style=pair_small_label_style),
                                                    html.Div(),
                                                    dcc.Input(id={"type": "sg-radius-bl", "level": level_str},
                                                              type="number", min=0, max=50, step="any", value=r_bl,
                                                              debounce=False, style=pair_input_style),
                                                    html.Div(),
                                                    html.Span("BR", style=pair_small_label_style),
                                                    html.Div(),
                                                    dcc.Input(id={"type": "sg-radius-br", "level": level_str},
                                                              type="number", min=0, max=50, step="any", value=r_br,
                                                              debounce=False, style=pair_input_style),
                                                ],
                                            ),
                                        ],
                                    ),
                                ],
                            ),
                        ],
                    ),

                    html.Div(
                        style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                               "justifyContent": "flex-start", "gap": "10px", "width": "100%"},
                        children=[
                            html.Label("Border width",
                                       style={**ctrl_label_style, "marginBottom": "0px", "width": right_label_w}),
                            html.Div(
                                style=right_ctrl_col_style,
                                children=[
                                    html.Div(
                                        style={"flex": "1", "minWidth": "0"},
                                        children=[
                                            dcc.Slider(id={"type": "sg-border-width-slider", "level": level_str},
                                                       min=0.0, max=3.0, step=0.05, value=border_width_v,
                                                       updatemode="mouseup", tooltip={"always_visible": False}),
                                        ],
                                    ),
                                    dcc.Input(debounce=True, id={"type": "sg-border-width-input", "level": level_str},
                                              type="number", min=0.0, max=5.0, step="any", value=border_width_v,
                                              style={"width": "100px", **ctrl_box_style}),
                                ],
                            ),
                        ],
                    ),

                    html.Div(
                        style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                               "justifyContent": "flex-start", "gap": "10px", "width": "100%"},
                        children=[
                            html.Label("Line style",
                                       style={**ctrl_label_style, "marginBottom": "0px", "width": right_label_w}),
                            html.Div(
                                style=right_ctrl_single_style,
                                children=[
                                    dcc.Dropdown(
                                        id={"type": "sg-line-style", "level": level_str},
                                        options=[
                                            {"label": "Solid", "value": "solid"},
                                            {"label": "Dash", "value": "dash"},
                                            {"label": "Dot", "value": "dot"},
                                            {"label": "Dashdot", "value": "dashdot"},
                                        ],
                                        value=line_style_v,
                                        clearable=False,
                                        optionHeight=dropdown_option_h,
                                        style={**ctrl_box_style, "width": "100%"},
                                    ),
                                ],
                            ),
                        ],
                    ),

                    html.Div(
                        style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                               "justifyContent": "flex-start", "gap": "10px", "width": "100%"},
                        children=[
                            html.Label("Border color",
                                       style={**ctrl_label_style, "marginBottom": "0px", "width": right_label_w}),
                            html.Div(
                                style=right_ctrl_col_style,
                                children=[
                                    dcc.Input(debounce=True, id={"type": "sg-border-color-picker", "level": level_str},
                                              type="color", value=border_color_v,
                                              style={"width": "48px", "height": "48px"}),
                                    dcc.Input(id={"type": "sg-border-color-text", "level": level_str}, type="text",
                                              value=border_color_v, debounce=True,
                                              style={"flex": "1", "minWidth": "0", **ctrl_box_style}),
                                ],
                            ),
                        ],
                    ),

                    html.Div(
                        style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                               "justifyContent": "flex-start", "gap": "10px", "width": "100%"},
                        children=[
                            html.Label("Border opacity",
                                       style={**ctrl_label_style, "marginBottom": "0px", "width": right_label_w}),
                            html.Div(
                                style=right_ctrl_col_style,
                                children=[
                                    html.Div(
                                        style={"flex": "1", "minWidth": "0"},
                                        children=[
                                            dcc.Slider(id={"type": "sg-border-opacity-slider", "level": level_str},
                                                       min=0.0, max=1.0, step=0.01, value=border_opacity_v,
                                                       updatemode="mouseup", tooltip={"always_visible": False}),
                                        ],
                                    ),
                                    dcc.Input(debounce=True, id={"type": "sg-border-opacity-input", "level": level_str},
                                              type="number", min=0.0, max=1.0, step="any", value=border_opacity_v,
                                              style={"width": "100px", **ctrl_box_style}),
                                ],
                            ),
                        ],
                    ),

                    html.Div(
                        style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                               "justifyContent": "flex-start", "gap": "10px", "width": "100%"},
                        children=[
                            html.Label("Padding",
                                       style={**ctrl_label_style, "marginBottom": "0px", "width": right_label_w}),
                            html.Div(
                                style=right_ctrl_single_style,
                                children=[
                                    html.Div(
                                        style=pair_two_input_row_style,
                                        children=[
                                            html.Span("L", style=pair_small_label_style),
                                            html.Div(),
                                            dcc.Input(debounce=True, id={"type": "sg-pad-left", "level": level_str},
                                                      type="number", min=0.0, max=5.0, step="any", value=pad_left_v,
                                                      style=pair_input_style),
                                            html.Div(),
                                            html.Span("R", style=pair_small_label_style),
                                            html.Div(),
                                            dcc.Input(debounce=True, id={"type": "sg-pad-right", "level": level_str},
                                                      type="number", min=0.0, max=5.0, step="any", value=pad_right_v,
                                                      style=pair_input_style),
                                        ],
                                    ),
                                ],
                            ),
                        ],
                    ),

                    html.Div(
                        style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                               "justifyContent": "flex-start", "gap": "10px", "width": "100%"},
                        children=[
                            html.Label("", style={**ctrl_label_style, "marginBottom": "0px", "width": right_label_w}),
                            html.Div(
                                style=right_ctrl_single_style,
                                children=[
                                    html.Div(
                                        style=pair_two_input_row_style,
                                        children=[
                                            html.Span("T", style=pair_small_label_style),
                                            html.Div(),
                                            dcc.Input(debounce=True, id={"type": "sg-pad-top", "level": level_str},
                                                      type="number", min=0.0, max=5.0, step="any", value=pad_top_v,
                                                      style=pair_input_style),
                                            html.Div(),
                                            html.Span("B", style=pair_small_label_style),
                                            html.Div(),
                                            dcc.Input(debounce=True, id={"type": "sg-pad-bottom", "level": level_str},
                                                      type="number", min=0.0, max=5.0, step="any", value=pad_bottom_v,
                                                      style=pair_input_style),
                                        ],
                                    ),
                                ],
                            ),
                        ],
                    ),

                    html.Div(
                        style={"boxSizing": "border-box"},
                        children=[
                            html.Button(btn_text, id={"type": "btn-sg-label-toggle", "level": level_str}, n_clicks=0,
                                        style={"width": "100%"}),
                        ],
                    ),

                    html.Div(
                        style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                               "justifyContent": "flex-start", "gap": "10px", "width": "100%"},
                        children=[
                            html.Label("Name font",
                                       style={**ctrl_label_style, "marginBottom": "0px", "width": right_label_w}),
                            html.Div(
                                style=right_ctrl_single_style,
                                children=[
                                    dcc.Dropdown(
                                        id={"type": "sg-label-font", "level": level_str},
                                        options=[
                                            {"label": "Arial", "value": "Arial"},
                                            {"label": "Times New Roman", "value": "Times New Roman"},
                                            {"label": "Helvetica", "value": "Helvetica"},
                                            {"label": "Courier New", "value": "Courier New"},
                                            {"label": "Georgia", "value": "Georgia"},
                                        ],
                                        value=label_font_v,
                                        clearable=False,
                                        optionHeight=dropdown_option_h,
                                        style={**ctrl_box_style, "width": "100%"},
                                    ),
                                ],
                            ),
                        ],
                    ),

                    html.Div(
                        style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                               "justifyContent": "flex-start", "gap": "10px", "width": "100%"},
                        children=[
                            html.Label("Name size",
                                       style={**ctrl_label_style, "marginBottom": "0px", "width": right_label_w}),
                            html.Div(
                                style=right_ctrl_col_style,
                                children=[
                                    html.Div(
                                        style={"flex": "1", "minWidth": "0"},
                                        children=[
                                            dcc.Slider(id={"type": "sg-label-size-slider", "level": level_str}, min=4,
                                                       max=24, step=1, value=label_size_v, updatemode="mouseup",
                                                       tooltip={"always_visible": False}),
                                        ],
                                    ),
                                    dcc.Input(debounce=True, id={"type": "sg-label-size-input", "level": level_str},
                                              type="number", min=6, max=48, step="any", value=label_size_v,
                                              style={"width": "100px", **ctrl_box_style}),
                                ],
                            ),
                        ],
                    ),

                    html.Div(
                        style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                               "justifyContent": "flex-start", "gap": "10px", "width": "100%"},
                        children=[
                            html.Label("Name color",
                                       style={**ctrl_label_style, "marginBottom": "0px", "width": right_label_w}),
                            html.Div(
                                style=right_ctrl_col_style,
                                children=[
                                    dcc.Input(debounce=True, id={"type": "sg-label-color-picker", "level": level_str},
                                              type="color", value=label_color_v,
                                              style={"width": "48px", "height": "48px"}),
                                    dcc.Input(id={"type": "sg-label-color-text", "level": level_str}, type="text",
                                              value=label_color_v, debounce=True,
                                              style={"flex": "1", "minWidth": "0", **ctrl_box_style}),
                                ],
                            ),
                        ],
                    ),

                    html.Div(
                        style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                               "justifyContent": "flex-start", "gap": "10px", "width": "100%"},
                        children=[
                            html.Label("Distance",
                                       style={**ctrl_label_style, "marginBottom": "0px", "width": right_label_w}),
                            html.Div(
                                style=right_ctrl_col_style,
                                children=[
                                    html.Div(
                                        style={"flex": "1", "minWidth": "0"},
                                        children=[
                                            dcc.Slider(id={"type": "sg-label-distance-slider", "level": level_str},
                                                       min=-2.0, max=2.0, step=0.01, value=label_distance_v,
                                                       updatemode="mouseup", tooltip={"always_visible": False}),
                                        ],
                                    ),
                                    dcc.Input(debounce=True, id={"type": "sg-label-distance-input", "level": level_str},
                                              type="number", min=-2.0, max=2.0, step="any", value=label_distance_v,
                                              style={"width": "100px", **ctrl_box_style}),
                                ],
                            ),
                        ],
                    ),

                    html.Div(
                        style={"display": "flex", "flexDirection": "row", "alignItems": "center",
                               "justifyContent": "flex-start", "gap": "10px", "width": "100%"},
                        children=[
                            html.Label("Name rotation",
                                       style={**ctrl_label_style, "marginBottom": "0px", "width": right_label_w}),
                            html.Div(
                                style=right_ctrl_col_style,
                                children=[
                                    html.Div(
                                        style={"flex": "1", "minWidth": "0"},
                                        children=[
                                            dcc.Slider(id={"type": "sg-text-rotation-slider", "level": level_str},
                                                       min=-180, max=180, step=1, value=label_rotation_v,
                                                       updatemode="mouseup", tooltip={"always_visible": False}),
                                        ],
                                    ),
                                    dcc.Input(debounce=True, id={"type": "sg-text-rotation", "level": level_str},
                                              type="number", min=-180, max=180, step="any", value=label_rotation_v,
                                              style={"width": "100px", **ctrl_box_style}),
                                ],
                            ),
                        ],
                    )
                ],
            ),
        ], style={"padding": "10px", "border": "2px solid #AAAAAA", "borderRadius": "6px", "marginBottom": "10px",
                  "backgroundColor": "#FAFAFA", "overflow": "hidden", "boxSizing": "border-box"})

    # Dynamic callback to generate supergroup panels based on metadata
    @app.callback(
        Output("supergroup-panels-container", "children"),
        Input("store-meta", "data"),
        State("store-level-labels", "data"),
        State("store-enclosure-visibility", "data"),
        State("store-style", "data"),
    )
    def render_supergroup_panels(meta, level_state, enclosure_visibility, style_state):
        levels = list((meta or {}).get("enclosure_levels", []) or [])
        show_map = dict((level_state or {}).get("show", {}) or {})
        enclosure_vis_map = dict(enclosure_visibility or {})

        # Count supergroup levels (exclude "group")
        sg_count = 0
        for lv in levels:
            key = str(lv.get("key", ""))
            if key.startswith("supergroup"):
                try:
                    num = int(key.replace("supergroup", ""))
                    if num > sg_count:
                        sg_count = num
                except Exception:
                    pass

        if sg_count == 0:
            return []

        supergroups_state = (style_state or {}).get("supergroups", {}) or {}

        panels = []
        for level in range(1, sg_count + 1):
            level_key = str(level)
            s = supergroups_state.get(level_key, _get_default_supergroup_style(level))
            is_shown = bool(show_map.get(f"supergroup{level}", False))
            # Default to visible (True) if not in the map
            is_enclosure_visible = enclosure_vis_map.get(level_key, True)
            panels.append(
                _render_supergroup_panel(level, s, ctrl_label_style, ctrl_box_style, is_shown, is_enclosure_visible,
                                         is_highest_level=(level == sg_count)))

        return panels

    # Callback to toggle enclosure visibility
    @app.callback(
        Output("store-enclosure-visibility", "data"),
        Input({"type": "btn-sg-visibility-toggle", "level": ALL}, "n_clicks"),
        State("store-enclosure-visibility", "data"),
        State("store-meta", "data"),
        prevent_initial_call=True,
    )
    def toggle_enclosure_visibility(n_clicks_list, visibility_state, meta):
        ctx = dash.callback_context
        if not ctx.triggered:
            raise PreventUpdate

        visibility_state = dict(visibility_state or {})

        # Get the triggered button
        trig = ctx.triggered[0]["prop_id"]
        trig_value = ctx.triggered[0].get("value")

        if not trig.endswith(".n_clicks"):
            raise PreventUpdate

        # Only process if there was an actual click (n_clicks > 0)
        # This prevents false triggers when panels are re-rendered
        if trig_value is None or trig_value <= 0:
            raise PreventUpdate

        try:
            btn_id_json = trig.split(".n_clicks")[0]
            btn_id = json.loads(btn_id_json)
            if btn_id.get("type") == "btn-sg-visibility-toggle":
                level_str = str(btn_id.get("level", ""))
                # Toggle visibility (default True if not set)
                current = visibility_state.get(level_str, True)
                visibility_state[level_str] = not current
        except Exception:
            pass

        return visibility_state

    # Update button text for sg name toggle (without re-rendering panels)
    @app.callback(
        Output({"type": "btn-sg-label-toggle", "level": MATCH}, "children"),
        Input("store-level-labels", "data"),
        State({"type": "btn-sg-label-toggle", "level": MATCH}, "id"),
        prevent_initial_call=True,
    )
    def update_sg_label_btn_text(level_state, btn_id):
        show_map = dict((level_state or {}).get("show", {}) or {})
        level = str(btn_id.get("level", ""))
        k = f"supergroup{level}"
        if bool(show_map.get(k, False)):
            return f"Hide enclosure{level} name"
        return f"Show enclosure{level} name"

    # Update button text for sg visibility toggle (without re-rendering panels)
    @app.callback(
        Output({"type": "btn-sg-visibility-toggle", "level": MATCH}, "children"),
        Input("store-enclosure-visibility", "data"),
        State({"type": "btn-sg-visibility-toggle", "level": MATCH}, "id"),
        prevent_initial_call=True,
    )
    def update_sg_vis_btn_text(visibility_state, btn_id):
        level = str(btn_id.get("level", ""))
        is_visible = (visibility_state or {}).get(level, True)
        if is_visible:
            return f"Hide enclosure{level}"
        return f"Show enclosure{level}"

        # Sync callbacks for Supergroup controls (per-level)

    @app.callback(
        Output({"type": "sg-border-width-slider", "level": MATCH}, "value"),
        Output({"type": "sg-border-width-input", "level": MATCH}, "value"),
        Input({"type": "sg-border-width-slider", "level": MATCH}, "value"),
        Input({"type": "sg-border-width-input", "level": MATCH}, "value"),
        prevent_initial_call=True,
    )
    def sync_sg_border_width(slider_v, input_v):
        ctx = dash.callback_context
        tid = getattr(ctx, "triggered_id", None)
        v = slider_v
        if isinstance(tid, dict) and tid.get("type") == "sg-border-width-input":
            v = input_v
        try:
            v2 = float(v)
        except Exception:
            v2 = float(slider_v) if slider_v is not None else 1.0
        v2 = max(0.0, min(5.0, v2))
        return v2, v2

    @app.callback(
        Output({"type": "sg-border-opacity-slider", "level": MATCH}, "value"),
        Output({"type": "sg-border-opacity-input", "level": MATCH}, "value"),
        Input({"type": "sg-border-opacity-slider", "level": MATCH}, "value"),
        Input({"type": "sg-border-opacity-input", "level": MATCH}, "value"),
        prevent_initial_call=True,
    )
    def sync_sg_border_opacity(slider_v, input_v):
        ctx = dash.callback_context
        tid = getattr(ctx, "triggered_id", None)
        v = slider_v
        if isinstance(tid, dict) and tid.get("type") == "sg-border-opacity-input":
            v = input_v
        try:
            v2 = float(v)
        except Exception:
            v2 = float(slider_v) if slider_v is not None else 0.35
        v2 = max(0.0, min(1.0, v2))
        return v2, v2

    @app.callback(
        Output({"type": "sg-label-size-slider", "level": MATCH}, "value"),
        Output({"type": "sg-label-size-input", "level": MATCH}, "value"),
        Input({"type": "sg-label-size-slider", "level": MATCH}, "value"),
        Input({"type": "sg-label-size-input", "level": MATCH}, "value"),
        prevent_initial_call=True,
    )
    def sync_sg_label_size(slider_v, input_v):
        ctx = dash.callback_context
        tid = getattr(ctx, "triggered_id", None)
        v = slider_v
        if isinstance(tid, dict) and tid.get("type") == "sg-label-size-input":
            v = input_v
        try:
            v2 = int(v)
        except Exception:
            v2 = int(slider_v) if slider_v is not None else 12
        v2 = max(6, min(48, v2))
        return v2, v2

    @app.callback(
        Output({"type": "sg-label-distance-slider", "level": MATCH}, "value"),
        Output({"type": "sg-label-distance-input", "level": MATCH}, "value"),
        Input({"type": "sg-label-distance-slider", "level": MATCH}, "value"),
        Input({"type": "sg-label-distance-input", "level": MATCH}, "value"),
        prevent_initial_call=True,
    )
    def sync_sg_label_distance(slider_v, input_v):
        ctx = dash.callback_context
        tid = getattr(ctx, "triggered_id", None)
        v = slider_v
        if isinstance(tid, dict) and tid.get("type") == "sg-label-distance-input":
            v = input_v
        try:
            v2 = float(v)
        except Exception:
            v2 = float(slider_v) if slider_v is not None else 0.08
        v2 = max(-2.0, min(2.0, v2))
        return v2, v2

    @app.callback(
        Output({"type": "sg-border-color-text", "level": MATCH}, "value"),
        Output({"type": "sg-border-color-picker", "level": MATCH}, "value"),
        Input({"type": "sg-border-color-text", "level": MATCH}, "value"),
        Input({"type": "sg-border-color-picker", "level": MATCH}, "value"),
        prevent_initial_call=True,
    )
    def sync_sg_border_color(text_v, picker_v):
        ctx = dash.callback_context
        tid = getattr(ctx, "triggered_id", None)
        v = text_v
        if isinstance(tid, dict) and tid.get("type") == "sg-border-color-picker":
            v = picker_v
        v2 = _normalize_hex_color(v, "#000000")
        return v2, v2

    @app.callback(
        Output({"type": "sg-label-color-text", "level": MATCH}, "value"),
        Output({"type": "sg-label-color-picker", "level": MATCH}, "value"),
        Input({"type": "sg-label-color-text", "level": MATCH}, "value"),
        Input({"type": "sg-label-color-picker", "level": MATCH}, "value"),
        prevent_initial_call=True,
    )
    def sync_sg_label_color(text_v, picker_v):
        ctx = dash.callback_context
        tid = getattr(ctx, "triggered_id", None)
        v = text_v
        if isinstance(tid, dict) and tid.get("type") == "sg-label-color-picker":
            v = picker_v
        v2 = _normalize_hex_color(v, "#000000")
        return v2, v2

    # Sync supergroup name rotation slider and input (pattern-matching)
    @app.callback(
        Output({"type": "sg-enclosure-gap-slider", "level": MATCH}, "value"),
        Output({"type": "sg-enclosure-gap-input", "level": MATCH}, "value"),
        Input({"type": "sg-enclosure-gap-slider", "level": MATCH}, "value"),
        Input({"type": "sg-enclosure-gap-input", "level": MATCH}, "value"),
        prevent_initial_call=True,
    )
    def sync_sg_enclosure_gap(slider_v, input_v):
        ctx = dash.callback_context
        tid = getattr(ctx, "triggered_id", None)
        v = slider_v
        if isinstance(tid, dict) and tid.get("type") == "sg-enclosure-gap-input":
            v = input_v
        try:
            v2 = float(v)
        except Exception:
            v2 = 2.0
        v2 = max(0.0, min(10.0, v2))
        return v2, v2

    # Sync supergroup name rotation slider and input (pattern-matching)
    @app.callback(
        Output({"type": "sg-text-rotation", "level": MATCH}, "value"),
        Output({"type": "sg-text-rotation-slider", "level": MATCH}, "value"),
        Input({"type": "sg-text-rotation", "level": MATCH}, "value"),
        Input({"type": "sg-text-rotation-slider", "level": MATCH}, "value"),
        prevent_initial_call=True,
    )
    def sync_sg_text_rotation(v_in, v_slider):
        ctx = dash.callback_context
        tid = getattr(ctx, "triggered_id", None)
        v = v_in if isinstance(tid, dict) and tid.get("type") == "sg-text-rotation" else v_slider
        if v is None:
            v = 0
        try:
            v = int(v)
        except Exception:
            v = 0
        v = max(-180, min(180, v))
        return v, v

    # Update store-style with supergroup settings (pattern-matching callback)
    @app.callback(
        Output("store-style", "data", allow_duplicate=True),
        Input({"type": "sg-pad-left", "level": ALL}, "value"),
        Input({"type": "sg-pad-right", "level": ALL}, "value"),
        Input({"type": "sg-pad-top", "level": ALL}, "value"),
        Input({"type": "sg-pad-bottom", "level": ALL}, "value"),
        Input({"type": "sg-radius-tl", "level": ALL}, "value"),
        Input({"type": "sg-radius-tr", "level": ALL}, "value"),
        Input({"type": "sg-radius-bl", "level": ALL}, "value"),
        Input({"type": "sg-radius-br", "level": ALL}, "value"),
        Input({"type": "sg-border-width-slider", "level": ALL}, "value"),
        Input({"type": "sg-border-width-input", "level": ALL}, "value"),
        Input({"type": "sg-border-opacity-slider", "level": ALL}, "value"),
        Input({"type": "sg-border-opacity-input", "level": ALL}, "value"),
        Input({"type": "sg-border-color-text", "level": ALL}, "value"),
        Input({"type": "sg-border-color-picker", "level": ALL}, "value"),
        Input({"type": "sg-line-style", "level": ALL}, "value"),
        Input({"type": "sg-label-font", "level": ALL}, "value"),
        Input({"type": "sg-label-size-slider", "level": ALL}, "value"),
        Input({"type": "sg-label-size-input", "level": ALL}, "value"),
        Input({"type": "sg-label-color-text", "level": ALL}, "value"),
        Input({"type": "sg-label-color-picker", "level": ALL}, "value"),
        Input({"type": "sg-label-distance-slider", "level": ALL}, "value"),
        Input({"type": "sg-label-distance-input", "level": ALL}, "value"),
        Input({"type": "sg-text-rotation", "level": ALL}, "value"),
        Input({"type": "sg-enclosure-gap-slider", "level": ALL}, "value"),
        Input({"type": "sg-enclosure-gap-input", "level": ALL}, "value"),
        State("store-style", "data"),
        prevent_initial_call=True,
    )
    def update_style_store_supergroups(
            _pad_lefts, _pad_rights, _pad_tops, _pad_bottoms,
            _r_tls, _r_trs, _r_bls, _r_brs,
            _bw_sliders, _bw_inputs,
            _bo_sliders, _bo_inputs,
            _bc_texts, _bc_pickers,
            _line_styles,
            _label_fonts,
            _ls_sliders, _ls_inputs,
            _lc_texts, _lc_pickers,
            _ld_sliders, _ld_inputs,
            _text_rotations,
            _eg_sliders, _eg_inputs,
            style_state,
    ):
        style_state = style_state or {}
        supergroups_state = dict(style_state.get("supergroups", {}) or {})

        ctx = dash.callback_context
        if not ctx.triggered:
            return style_state

        trig = ctx.triggered[0]
        trig_id = trig.get("prop_id", "")
        trig_value = trig.get("value", None)

        if not trig_id.startswith("{"):
            return style_state

        try:
            id_json = trig_id.rsplit(".", 1)[0]
            id_dict = json.loads(id_json)
            level_str = str(id_dict.get("level", ""))
            input_type = str(id_dict.get("type", ""))
        except Exception:
            return style_state

        if not level_str:
            return style_state

        if level_str not in supergroups_state:
            supergroups_state[level_str] = _get_default_supergroup_style(int(level_str))

        level_state = dict(supergroups_state[level_str])

        def _safe_float(v, default, min_v=0.0, max_v=2.0):
            # Handle None default
            if default is None:
                default = min_v
            try:
                if v is None or v == "" or v == "None":
                    return float(default)
                return float(max(min_v, min(max_v, float(v))))
            except (ValueError, TypeError):
                return float(default)

        if input_type == "sg-pad-left":
            level_state["pad_left"] = _safe_float(trig_value, level_state.get("pad_left", 0.0), 0.0, 5.0)
        elif input_type == "sg-pad-right":
            level_state["pad_right"] = _safe_float(trig_value, level_state.get("pad_right", 0.0), 0.0, 5.0)
        elif input_type == "sg-pad-top":
            level_state["pad_top"] = _safe_float(trig_value, level_state.get("pad_top", 0.0), 0.0, 5.0)
        elif input_type == "sg-pad-bottom":
            level_state["pad_bottom"] = _safe_float(trig_value, level_state.get("pad_bottom", 0.0), 0.0, 5.0)
        elif input_type == "sg-radius-tl":
            r_raw = level_state.get("radii")
            r = _safe_radii_list(r_raw)
            r[0] = _clamp_radius(trig_value)
            level_state["radii"] = r
        elif input_type == "sg-radius-tr":
            r_raw = level_state.get("radii")
            r = _safe_radii_list(r_raw)
            r[1] = _clamp_radius(trig_value)
            level_state["radii"] = r
        elif input_type == "sg-radius-bl":
            r_raw = level_state.get("radii")
            r = _safe_radii_list(r_raw)
            r[3] = _clamp_radius(trig_value)
            level_state["radii"] = r
        elif input_type == "sg-radius-br":
            r_raw = level_state.get("radii")
            r = _safe_radii_list(r_raw)
            r[2] = _clamp_radius(trig_value)
            level_state["radii"] = r
        elif input_type in ("sg-border-width-slider", "sg-border-width-input"):
            level_state["border_width"] = _safe_float(trig_value, level_state.get("border_width", 1.0), 0.0, 5.0)
        elif input_type in ("sg-border-opacity-slider", "sg-border-opacity-input"):
            level_state["border_opacity"] = _safe_float(trig_value, level_state.get("border_opacity",
                                                                                    _get_default_supergroup_style(
                                                                                        int(level_str)).get(
                                                                                        "border_opacity", 0.35)), 0.0,
                                                        1.0)
        elif input_type in ("sg-border-color-text", "sg-border-color-picker"):
            level_state["border_color"] = _normalize_hex_color(trig_value, "#000000")
        elif input_type == "sg-line-style":
            level_state["line_style"] = str(trig_value or "solid")
        elif input_type == "sg-label-font":
            level_state["label_font"] = str(trig_value or "Arial")
        elif input_type in ("sg-label-size-slider", "sg-label-size-input"):
            try:
                level_state["label_size"] = int(max(6, min(48, int(trig_value))))
            except Exception:
                level_state["label_size"] = int(level_state.get("label_size", 23))
        elif input_type in ("sg-label-color-text", "sg-label-color-picker"):
            level_state["label_color"] = _normalize_hex_color(trig_value, "#000000")
        elif input_type in ("sg-label-distance-slider", "sg-label-distance-input"):
            level_state["label_distance"] = _safe_float(trig_value, level_state.get("label_distance", -0.05), -2.0, 2.0)
        elif input_type == "sg-text-rotation":
            try:
                rot = int(trig_value)
                rot = max(-180, min(180, rot))
            except Exception:
                rot = 0
            level_state["label_rotation"] = rot
        elif input_type in ("sg-enclosure-gap-slider", "sg-enclosure-gap-input"):
            level_state["enclosure_gap"] = _safe_float(trig_value, level_state.get("enclosure_gap", 2.0), 0.0, 10.0)

        supergroups_state[level_str] = level_state
        style_state["supergroups"] = supergroups_state
        return style_state

    @app.callback(
        Output("store-base-fig", "data"),
        Output("store-meta", "data"),
        Output("store-ilog-event", "data", allow_duplicate=True),
        Input("category-col", "value"),
        Input("store-style", "data"),
        Input("store-upload-slices", "data"),
        Input("store-band-mode", "data"),
        Input("store-block-mode", "data"),
        Input("store-collapse-trigger", "data"),
        Input("store-sweep-params", "data"),
        # graph_size is a STATE, not an Input. Watching it as Input would
        # make every figure-render → graph-size remeasure → rebuild_base
        # cycle into an infinite loop, which manifested as the loading
        # overlay being triggered on every action (each rebuild took long
        # enough to cross the 550ms CSS delay-before-show threshold).
        # The trade-off: the very first render uses the default 600×700
        # canvas size for text wrapping. The subsequent measurements update
        # store-graph-size silently; the next user-driven rebuild (Generate /
        # style change / etc.) will pick up the real size.
        State("store-graph-size", "data"),
        State("store-collapse", "data"),
        State("store-selected", "data"),
        State("store-meta", "data"),
        State("url", "search"),
        prevent_initial_call=True,
        running=[
            (Output("store-loading-rebuild", "data"), True, False),
        ],
    )
    def rebuild_base(category_col_value, style_state, upload_state, band_mode_state, block_mode_state,
                     _collapse_trigger, sweep_params, graph_size,
                     collapse_state, _selected_state, prev_meta, url_search):
        import time as _t
        t0 = _t.time()
        paths = list((upload_state or {}).get("paths", []) or [])
        x_idx = 0
        y_idx = 1 if len(paths) > 1 else 0
        band_mode = (band_mode_state or {}).get("mode", "existence")
        band_proportion = float((band_mode_state or {}).get("proportion", 0.5))
        block_mode = (block_mode_state or {}).get("mode", "existence")
        block_median_width = float((block_mode_state or {}).get("median_width", cfg.block_width))
        sweep_params = sweep_params or {}
        k_max_use = int(sweep_params.get("k_max", 10))
        m_use = int(sweep_params.get("m", 2))
        delta_use = float(sweep_params.get("delta", 0.01))
        # On collapse, preserve the y_range from the pre-collapse figure so the
        # diagram collapses in place without rescaling.
        initial_y_range = None
        if (collapse_state or {}).get("collapsed"):
            initial_y_range = (prev_meta or {}).get("initial_y_range")
        fig, meta = build_base(category_col_value, x_idx, y_idx, style_state=style_state, upload_state=upload_state,
                               band_mode=band_mode, band_proportion=band_proportion, block_mode=block_mode,
                               block_median_width=block_median_width, collapse_state=collapse_state,
                               sweep_k_max=k_max_use, sweep_m=m_use, sweep_delta=delta_use,
                               graph_size=graph_size, initial_y_range=initial_y_range)
        # ilog: only log if category-col triggered this rebuild
        ctx = dash.callback_context
        trig_id = ctx.triggered[0].get("prop_id", "") if ctx.triggered else ""
        ilog_evt = dash.no_update
        if "category-col" in trig_id:
            _params = {}
            if url_search:
                for _p in url_search.lstrip("?").split("&"):
                    if "=" in _p:
                        _k, _v = _p.split("=", 1)
                        _params[_k] = _v
            ilog_evt = {"type": "category_select", "event_num": 9,
                        "value": category_col_value,
                        "ts": int(_t.time() * 1000),
                        "pid": _params.get("pid", ""), "qid": _params.get("qid", ""),
                        "condition": "A"}
        
        t_rebuild = _t.time() - t0
        print(f"DEBUG: rebuild_base took {t_rebuild:.3f}s (trig by {trig_id})")
        
        return fig.to_dict(), meta, ilog_evt

    # Combined band mode and proportion callback - avoids circular dependency
    @app.callback(
        Output("store-band-mode", "data"),
        Output("btn-band-mode", "children"),
        Output("band-width-label", "children"),
        Output("band-width-slider", "min"),
        Output("band-width-slider", "max"),
        Output("band-width-slider", "value"),
        Output("band-width-input", "min"),
        Output("band-width-input", "max"),
        Output("band-width-input", "value"),
        Input("btn-band-mode", "n_clicks"),
        Input("band-width-slider", "value"),
        Input("band-width-input", "value"),
        State("store-band-mode", "data"),
        State("store-style", "data"),
    )
    def update_band_mode_and_controls(n_clicks, slider_val, input_val, band_mode_state, style_state):
        ctx = dash.callback_context
        trig = (ctx.triggered[0]["prop_id"] if (ctx.triggered and len(ctx.triggered) > 0) else "")

        current_mode = (band_mode_state or {}).get("mode", "existence")
        current_proportion = float((band_mode_state or {}).get("proportion", 0.5))

        new_mode = current_mode
        new_proportion = current_proportion

        # Handle mode toggle button
        if trig.startswith("btn-band-mode"):
            new_mode = "strength" if current_mode == "existence" else "existence"

        # Prepare outputs
        btn_text = "Switch to Strength" if new_mode == "existence" else "Switch to Existence"

        if new_mode == "strength":
            # Strength mode: control proportion (0.1 ~ 1.0)
            label = "Proportion"
            min_val = 0.1
            max_val = 1.0

            # Handle proportion change
            if trig.startswith("band-width-slider") and slider_val is not None:
                new_proportion = float(max(0.1, min(1.0, float(slider_val))))
            elif trig.startswith("band-width-input") and input_val is not None:
                new_proportion = float(max(0.1, min(1.0, float(input_val))))

            value = new_proportion
        else:
            # Existence mode: control width ratio (0.02 ~ 0.30) — global for all types
            label = "Width"
            min_val = 0.02
            max_val = 0.30

            # Get width from style_state (all types share the same value now)
            band_state = ((style_state or {}).get("band", {}) or {})
            w_map = dict(band_state.get("width_ratios") or {})
            for k in cfg.band_colors.keys():
                w_map.setdefault(k, float(cfg.band_width_ratio))

            def _all_same():
                vals = list(w_map.values())
                return vals[0] if vals and all(abs(v - vals[0]) < 1e-9 for v in vals) else cfg.band_width_ratio

            # Handle different triggers
            if trig.startswith("band-width-input") and input_val is not None:
                value = float(input_val)
            elif trig.startswith("band-width-slider") and slider_val is not None:
                value = float(slider_val)
            else:
                value = _all_same()

            value = max(min_val, min(max_val, value))

        return (
            {"mode": new_mode, "proportion": new_proportion},
            btn_text,
            label,
            min_val, max_val, value,
            min_val, max_val, value,
        )

    # Combined block mode and median_width callback
    @app.callback(
        Output("store-block-mode", "data"),
        Output("btn-block-mode", "children"),
        Output("block-width-label", "children"),
        Output("block-width-slider", "min"),
        Output("block-width-slider", "max"),
        Output("block-width-slider", "value"),
        Output("block-width-input", "min"),
        Output("block-width-input", "max"),
        Output("block-width-input", "value"),
        Input("btn-block-mode", "n_clicks"),
        Input("block-width-slider", "value"),
        Input("block-width-input", "value"),
        Input("block-type", "data"),
        State("store-block-mode", "data"),
        State("store-style", "data"),
    )
    def update_block_mode_and_controls(n_clicks, slider_val, input_val, block_type, block_mode_state, style_state):
        ctx = dash.callback_context
        trig = (ctx.triggered[0]["prop_id"] if (ctx.triggered and len(ctx.triggered) > 0) else "")

        current_mode = (block_mode_state or {}).get("mode", "existence")
        current_median_width = float((block_mode_state or {}).get("median_width", cfg.block_width))

        new_mode = current_mode
        new_median_width = current_median_width

        # Handle mode toggle button
        if trig.startswith("btn-block-mode"):
            new_mode = "strength" if current_mode == "existence" else "existence"

        # Prepare outputs
        btn_text = "Switch to Strength" if new_mode == "existence" else "Switch to Existence"

        if new_mode == "strength":
            # Strength mode: control median width
            label = "Median Width"
            min_val = 0.5
            max_val = 5.0

            # Handle median width change
            if trig.startswith("block-width-slider") and slider_val is not None:
                new_median_width = float(max(0.5, min(5.0, float(slider_val))))
            elif trig.startswith("block-width-input") and input_val is not None:
                new_median_width = float(max(0.5, min(5.0, float(input_val))))

            value = new_median_width
        else:
            # Existence mode: control uniform width
            label = "Width"
            min_val = 0.5
            max_val = 5.0

            # Get width from style_state
            block_state = ((style_state or {}).get("block", {}) or {})
            widths = dict(block_state.get("widths") or {})
            for k in cfg.block_colors.keys():
                widths.setdefault(k, float(cfg.block_width))

            def _all_same():
                vals = list(widths.values())
                return vals[0] if vals and all(abs(v - vals[0]) < 1e-9 for v in vals) else float(cfg.block_width)

            # Handle different triggers
            if trig.startswith("block-type"):
                value = _all_same() if block_type == "All" else float(widths.get(block_type, float(cfg.block_width)))
            elif trig.startswith("block-width-input") and input_val is not None:
                value = float(input_val)
            elif trig.startswith("block-width-slider") and slider_val is not None:
                value = float(slider_val)
            else:
                # Default (including mode switch): get from widths based on current block_type
                if block_type == "All":
                    value = _all_same()
                else:
                    value = float(widths.get(block_type, float(cfg.block_width)))

            value = max(min_val, min(max_val, value))

        return (
            {"mode": new_mode, "median_width": new_median_width},
            btn_text,
            label,
            min_val, max_val, value,
            min_val, max_val, value,
        )

    @app.callback(
        Output("store-level-labels", "data"),
        Input("store-meta", "data"),
        Input("btn-group-label-toggle", "n_clicks"),
        Input({"type": "btn-sg-label-toggle", "level": ALL}, "n_clicks"),
        State("store-level-labels", "data"),
    )
    def update_level_label_state(meta, group_btn_clicks, sg_btn_clicks, level_state):
        level_state = level_state or {"show": {}, "last_category_col": None, "last_valid_keys": []}
        show_map = dict(level_state.get("show", {}) or {})
        last_category_col = level_state.get("last_category_col", None)
        last_valid_keys = set(level_state.get("last_valid_keys", []) or [])

        levels = list((meta or {}).get("enclosure_levels", []) or [])
        valid_keys = [str(lv.get("key", "")) for lv in levels if str(lv.get("key", "")) != ""]
        valid_keys_set = set(valid_keys)

        # Get current category_col from meta
        current_category_col = (meta or {}).get("category_col", None)

        ctx = dash.callback_context
        trig = (ctx.triggered[0]["prop_id"] if (ctx.triggered and len(ctx.triggered) > 0) else "")
        trig_value = ctx.triggered[0].get("value") if (ctx.triggered and len(ctx.triggered) > 0) else None

        # Check if triggered by ACTUAL button click (n_clicks > 0)
        # This prevents false triggers when panels are re-rendered
        is_group_button_click = (
                trig == "btn-group-label-toggle.n_clicks" and
                group_btn_clicks is not None and
                group_btn_clicks > 0
        )

        is_sg_button_click = False
        sg_clicked_level = None
        if trig.endswith(".n_clicks") and trig.startswith("{"):
            try:
                btn_id_json = trig.split(".", 1)[0]
                btn_id = json.loads(btn_id_json)
                if btn_id.get("type") == "btn-sg-label-toggle":
                    # Verify this button actually has clicks (not just re-rendered)
                    if trig_value is not None and trig_value > 0:
                        is_sg_button_click = True
                        sg_clicked_level = str(btn_id.get("level", ""))
            except Exception:
                pass

        is_button_trigger = is_group_button_click or is_sg_button_click

        # Check if category_col changed or valid_keys changed (structure changed)
        category_changed = (current_category_col != last_category_col)
        structure_changed = (valid_keys_set != last_valid_keys)

        # If meta changed (not button click) and either category or structure changed, reset all
        # Exception: on first load (last_category_col is None), preserve initial show_map (e.g. group=True)
        is_first_load = (last_category_col is None)
        if not is_button_trigger and (category_changed or structure_changed) and not is_first_load:
            # Reset all label visibility when category or structure changes
            # group defaults to True (shown), supergroups default to False (hidden)
            show_map = {k: (k == "group") for k in valid_keys}
        else:
            # Ensure keys exist with defaults: group=True (shown by default), supergroups=False
            for k in valid_keys:
                if k not in show_map:
                    show_map[k] = (k == "group")

            # Drop keys that are no longer present in the current metadata
            show_map = {k: bool(v) for k, v in show_map.items() if k in valid_keys_set}

        # Handle group label toggle button
        if is_group_button_click:
            show_map["group"] = not bool(show_map.get("group", False))

        # Handle supergroup label toggle buttons (pattern-matching)
        if is_sg_button_click and sg_clicked_level:
            k = f"supergroup{sg_clicked_level}"
            if k in show_map:
                show_map[k] = not bool(show_map.get(k, False))

        return {
            "show": show_map,
            "last_category_col": current_category_col,
            "last_valid_keys": list(valid_keys)
        }

    @app.callback(
        Output("btn-group-label-toggle", "children"),
        Input("store-level-labels", "data"),
    )
    def update_group_label_toggle_text(level_state):
        show_map = {}
        if isinstance(level_state, dict):
            show_map = dict(level_state.get("show", {}) or {})
        if bool(show_map.get("group", False)):
            return "Hide aggregation name"
        return "Show aggregation name"

    @app.callback(
        Output("btn-slice-toggle", "children", allow_duplicate=True),
        Input("store-slice-style", "data"),
        prevent_initial_call=True,
    )
    def update_slice_toggle_text(slice_style_state):
        st = slice_style_state or {}
        if bool(st.get("visible", True)):
            return "Hide slice labels"
        return "Show slice labels"

    # --- Slice rename: generate input rows with anchor checkboxes from meta ---
    @app.callback(
        Output("slice-rename-container", "children"),
        Input("store-meta", "data"),
        Input("store-collapse", "data"),
        State("store-slice-rename", "data"),
    )
    def generate_slice_rename_inputs(meta, collapse_state, rename_data):
        if not meta:
            raise PreventUpdate
        slices = meta.get("slices", [])
        if not slices:
            return []
        rename_data = rename_data or {}
        collapse_state = collapse_state or {}
        anchor_x = collapse_state.get("anchor_x")
        anchor_y = collapse_state.get("anchor_y")
        rows = []
        for i, sid in enumerate(slices):
            display_name = rename_data.get(str(sid), str(sid))
            is_checked = (anchor_x == i) or (anchor_y == i)
            rows.append(
                html.Div(
                    style={"display": "flex", "flexDirection": "row", "alignItems": "center", "gap": "10px",
                           "marginTop": "10px", "width": "100%"},
                    children=[
                        html.Label(f"Slice {i + 1}",
                                   style={"fontSize": "40px", "whiteSpace": "nowrap", "flexShrink": "0",
                                          "marginBottom": "0px", "width": "220px"}),
                        html.Div(
                            style={"display": "block", "flex": "1", "minWidth": "0", "boxSizing": "border-box"},
                            children=[
                                dcc.Input(
                                    id={"type": "slice-rename-input", "index": str(sid)},
                                    type="text",
                                    value=display_name,
                                    debounce=True,
                                    style={"width": "100%", "boxSizing": "border-box"},
                                ),
                            ],
                        ),
                        dcc.Checklist(
                            id={"type": "slice-anchor-check", "index": str(i)},
                            options=[{"label": "", "value": "on"}],
                            value=["on"] if is_checked else [],
                            style={"display": "flex", "alignItems": "center", "marginLeft": "6px"},
                            inputStyle={"width": "44px", "height": "44px", "cursor": "pointer"},
                        ),
                    ],
                )
            )
        return rows

    # --- Slice rename: collect input values into store ---
    @app.callback(
        Output("store-slice-rename", "data"),
        Input({"type": "slice-rename-input", "index": ALL}, "value"),
        State("store-meta", "data"),
        State("store-slice-rename", "data"),
        prevent_initial_call=True,
    )
    def collect_slice_renames(values, meta, rename_data):
        if not meta:
            raise PreventUpdate
        slices = meta.get("slices", [])
        if not slices:
            raise PreventUpdate
        # Use dash.callback_context.inputs_list for reliable sid→value mapping
        new_rename = {}
        cb_ctx = dash.callback_context
        inputs_list = cb_ctx.inputs_list
        if inputs_list and len(inputs_list) > 0:
            items = inputs_list[0] if isinstance(inputs_list[0], list) else inputs_list
            for item in items:
                if not isinstance(item, dict):
                    continue
                item_id = item.get("id", {})
                sid = str(item_id.get("index", "")) if isinstance(item_id, dict) else ""
                val = str(item.get("value", "") or "").strip()
                if sid and val and val != sid:
                    new_rename[sid] = val
        return new_rename

    # --- Collapse panel: Collapse/Expand buttons and anchor checkboxes ---
    @app.callback(
        Output("store-collapse", "data"),
        Output("store-collapse-trigger", "data"),
        Output("store-ilog-event", "data", allow_duplicate=True),
        Input("btn-collapse", "n_clicks"),
        Input("btn-expand", "n_clicks"),
        Input("btn-toggle-collapse-line", "n_clicks"),
        Input({"type": "slice-anchor-check", "index": ALL}, "value"),
        State("store-meta", "data"),
        State("store-collapse", "data"),
        State("store-collapse-trigger", "data"),
        State("url", "search"),
        prevent_initial_call=True,
    )
    def handle_collapse_panel(n_collapse, n_expand, n_toggle_line, check_values, meta, collapse_state, trigger_val, url_search):
        import time as _t
        collapse_state = collapse_state or {"anchor_x": None, "anchor_y": None, "collapsed": False,
                                            "line_hidden": False}
        new_state = dict(collapse_state)

        ctx = dash.callback_context
        if not ctx.triggered:
            raise PreventUpdate
        trig_id = ctx.triggered[0].get("prop_id", "")

        def _ilog(evt_type, extra=None):
            _params = {}
            if url_search:
                for _p in (url_search or "").lstrip("?").split("&"):
                    if "=" in _p:
                        _k, _v = _p.split("=", 1)
                        _params[_k] = _v
            ev = {"type": evt_type, "ts": int(_t.time() * 1000),
                  "pid": _params.get("pid", ""), "qid": _params.get("qid", ""),
                  "condition": "A"}
            if extra:
                ev.update(extra)
            return ev

        # --- Checkbox changes: update anchor_x / anchor_y ---
        if "slice-anchor-check" in trig_id:
            # Collect which indices are checked
            checked_indices = []
            cb_inputs = ctx.inputs_list
            # cb_inputs[3] corresponds to the slice-anchor-check pattern
            check_list = cb_inputs[3] if len(cb_inputs) > 3 else []
            if isinstance(check_list, list):
                for item in check_list:
                    if not isinstance(item, dict):
                        continue
                    item_id = item.get("id", {})
                    idx_str = str(item_id.get("index", "")) if isinstance(item_id, dict) else ""
                    val = item.get("value", [])
                    if val and "on" in val:
                        try:
                            checked_indices.append(int(idx_str))
                        except (ValueError, TypeError):
                            pass

            checked_indices.sort()

            # Only allow at most 2 selections
            if len(checked_indices) == 0:
                new_ax, new_ay = None, None
            elif len(checked_indices) == 1:
                new_ax, new_ay = checked_indices[0], None
            else:
                new_ax, new_ay = checked_indices[0], checked_indices[1]

            # Guard: if no actual change, prevent infinite loop from re-rendered checkboxes
            if new_ax == new_state.get("anchor_x") and new_ay == new_state.get("anchor_y"):
                raise PreventUpdate

            new_state["anchor_x"] = new_ax
            new_state["anchor_y"] = new_ay

            # 解析本次触发的 checkbox index 和方向
            import json as _json_ilog_cb
            toggled_slice = None
            toggled_action = None
            try:
                trig_id_str = trig_id.rsplit(".", 1)[0]
                trig_obj = _json_ilog_cb.loads(trig_id_str)
                toggled_idx = int(trig_obj.get("index", -1))
                toggled_slice = toggled_idx + 1  # 转为 1-based
                toggled_action = "select" if toggled_idx in checked_indices else "deselect"
            except Exception:
                pass

            return new_state, dash.no_update, _ilog("slice_label_check", {
                "event_num": 5,
                "toggled_slice": toggled_slice,
                "action": toggled_action,
                "checked_slices": [i + 1 for i in checked_indices],  # 1-based
            })

        _next_trigger = (trigger_val or 0) + 1

        # --- Collapse button ---
        if "btn-collapse" in trig_id:
            ax = new_state.get("anchor_x")
            ay = new_state.get("anchor_y")
            if ax is not None and ay is not None and ax != ay:
                # Ensure anchor_x < anchor_y
                if ax > ay:
                    ax, ay = ay, ax
                    new_state["anchor_x"] = ax
                    new_state["anchor_y"] = ay
                # Only collapse if there's at least one layer between anchors
                if ay - ax > 1:
                    new_state["collapsed"] = True
                    new_state["line_hidden"] = new_state.get("line_hidden", False)
            return new_state, _next_trigger, _ilog("collapse_click", {"event_num": 6})

        # --- Expand button ---
        if "btn-expand" in trig_id:
            new_state["collapsed"] = False
            new_state["anchor_x"] = None
            new_state["anchor_y"] = None
            return new_state, _next_trigger, _ilog("expand_click", {"event_num": 7})

        # --- Toggle collapse line visibility ---
        if "btn-toggle-collapse-line" in trig_id:
            new_state["line_hidden"] = not new_state.get("line_hidden", False)
            return new_state, _next_trigger, dash.no_update

        raise PreventUpdate

    # --- Toggle collapse line button label ---
    @app.callback(
        Output("btn-toggle-collapse-line", "children"),
        Input("store-collapse", "data"),
    )
    def update_collapse_line_btn_label(collapse_state):
        collapse_state = collapse_state or {}
        if collapse_state.get("line_hidden", False):
            return "Show collapse line"
        return "Hide collapse line"

    @app.callback(
        Output("graph", "figure"),
        Output("store-selected", "data"),
        Output("store-exo", "data"),
        Output("store-click-nonce", "data"),
        Output("store-bands", "data"),
        Output("btn-exo-all", "children"),
        Output("btn-bands-all", "children"),
        Output("store-ilog-event", "data", allow_duplicate=True),
        Input("graph", "clickData"),
        Input("graph", "selectedData"),
        Input("btn-wrap-leftclick", "n_clicks"),
        Input("btn-exo-all", "n_clicks"),
        Input("btn-bands-all", "n_clicks"),
        Input("store-bands", "data"),
        Input("store-base-fig", "data"),
        Input("store-meta", "data"),
        Input("store-level-labels", "data"),
        Input("store-slice-style", "data"),
        Input("store-enclosure-visibility", "data"),
        Input("store-slice-rename", "data"),
        State("store-zoom-state", "data"),
        State("store-selected", "data"),
        State("store-exo", "data"),
        State("store-click-nonce", "data"),
        State("store-style", "data"),
        State("store-graph-size", "data"),
        State("url", "search"),
        prevent_initial_call=True,  # graph already initialised with fig0; prevents store-bands circular on load
        running=[
            # Flips True the instant Dash queues this callback (client-side,
            # before any server work). Drives the top-bar status badge,
            # giving the user immediate feedback on click.
            (Output("store-loading-highlight", "data"), True, False),
        ],
    )
    def render_with_highlight(clickData, selectedData, _n_wrap, _n_all, _n_bands_all, bands_state_input, base_fig, meta,
                              level_state, slice_style_state, enclosure_visibility, slice_rename_data, zoom_state,
                              selected_state, exo_state, click_nonce_state, style_state, graph_size, url_search):
        import time as _t
        t0 = _t.time()
        if not base_fig or not meta:
            raise PreventUpdate

        curve_to_id = meta.get("curve_to_id", [])
        slices = meta.get("slices", [])
        
        ctx = dash.callback_context
        trig_props = [t.get("prop_id", "") for t in (ctx.triggered or [])]
        trig_set = set(trig_props)
        trig_main = trig_props[0] if trig_props else "unknown"

        exo_state = exo_state or {}
        hidden_slices = list(exo_state.get("hidden_slices", []))
        btn_nonce_map = dict(exo_state.get("nonce", {}))

        click_nonce_state = click_nonce_state or {}
        click_nonce_map = dict(click_nonce_state.get("nonce", {}))

        bands_state = bands_state_input or {"hidden_types": []}
        hidden_types = set([str(x) for x in (bands_state.get("hidden_types", []) or [])])
        hide_all_bands = ("__ALL__" in hidden_types)

        active_sel = (selected_state or {}).get("selected_id", None)

        def resolve_id(data):
            if not data or "points" not in data or not data["points"]:
                return None
            cn = data["points"][0].get("curveNumber", None)
            if cn is None:
                return None
            cn = int(cn)
            if cn < 0 or cn >= len(curve_to_id):
                return None
            oid = curve_to_id[cn]
            return oid.replace("::TEXT", "")

        trig_btn_all = any(p.startswith("btn-exo-all") for p in trig_set)
        trig_btn_bands_all = any(p.startswith("btn-bands-all") for p in trig_set)
        trig_clickdata = any(p.startswith("graph.clickData") for p in trig_set)
        trig_selected = any(p.startswith("graph.selectedData") for p in trig_set)
        trig_wrap = any(p.startswith("btn-wrap-leftclick.n_clicks") for p in trig_set)
        trig_new_data = any(p.startswith("store-meta") for p in trig_set)
        # On first load or new data, hide Exo blocks and Inflow/Outflow bands
        if (exo_state.get("hide_all_initial") and not hidden_slices and slices) or trig_new_data:
            hidden_slices = list(slices)
        if trig_new_data:
            hidden_types = {"Inflow", "Outflow"}

        # Helper: treat selectedData empty as deselect
        def is_deselect(sd) -> bool:
            if sd is None:
                return True
            pts = sd.get("points", None)
            return (pts is None) or (len(pts) == 0)

        # --- A) Top button: toggle all exo visibility (does NOT suppress per-slice toggles) ---
        if trig_btn_all:
            hidden_set = set(hidden_slices) & set(slices)
            if len(slices) > 0 and len(hidden_set) == len(slices):
                hidden_slices = []
            else:
                hidden_slices = list(slices)

            # bump nonce for all slices so next click on any left button always triggers
            for sid in slices:
                btn_nonce_map[sid] = int(btn_nonce_map.get(sid, 0)) + 1

        # --- A2) Global button: toggle all band visibility ---
        if trig_btn_bands_all:
            if len(hidden_types) > 0:
                hidden_types.clear()
            else:
                hidden_types.add("__ALL__")
            hide_all_bands = ("__ALL__" in hidden_types)

        # --- B) Graph clicks (highest priority) ---
        if trig_clickdata:
            click_id = resolve_id(clickData)
            # Ignore clicks on enclosure label traces
            if click_id and str(click_id).startswith("LBL::"):
                click_id = None

            # bump nonce for BG/BLOCK/BAND so repeated clicks still change clickData
            if click_id and (
                    click_id == "BG::CLICK" or click_id.startswith("BLOCK::") or click_id.startswith("BAND::")):
                click_nonce_map[click_id] = int(click_nonce_map.get(click_id, 0)) + 1

            if click_id == "BG::CLICK":
                active_sel = None

            elif click_id and click_id.startswith("BTNEXO::"):
                # Button exists only when Exo is hidden, so this means: SHOW Exo
                sid = click_id.split("::", 1)[1]
                if sid in hidden_slices:
                    hidden_slices = [s for s in hidden_slices if s != sid]

                # bump nonce for THIS slice (keep repeated clicks safe)
                btn_nonce_map[sid] = int(btn_nonce_map.get(sid, 0)) + 1

            elif click_id and click_id.startswith("BLOCK::"):
                # If user clicks Exo block: HIDE Exo (no highlight)
                parts = click_id.split("::")
                if len(parts) >= 3 and parts[2] == EXO_NAME:
                    sid = parts[1]
                    if sid not in hidden_slices:
                        hidden_slices = hidden_slices + [sid]
                    btn_nonce_map[sid] = int(btn_nonce_map.get(sid, 0)) + 1
                else:
                    # normal block highlight toggle
                    if active_sel == click_id:
                        active_sel = None
                    else:
                        active_sel = click_id

            elif click_id and click_id.startswith("BAND::"):
                # band highlight toggle
                if active_sel == click_id:
                    active_sel = None
                else:
                    active_sel = click_id

        # --- B2) Click on graph container blank (outside Plotly's click area) ---
        # BG::CLICK trace (background polygon) already handles all blank clicks
        # within the data range via trig_clickdata.  The btn-wrap path used to
        # clear active_sel here, but it fires as a SEPARATE Dash invocation
        # AFTER trig_clickdata, which caused the just-set selection to be
        # immediately cleared.  We only clear when BOTH fire together (same
        # invocation) and clickData resolved to nothing meaningful.
        if trig_wrap and trig_clickdata:
            # Both fired simultaneously: clickData took priority above; nothing extra needed.
            pass
        elif trig_wrap and (not trig_clickdata) and (not trig_btn_all):
            # btn-wrap fired alone (click was outside Plotly's plot area entirely).
            # Only clear if active_sel is not a freshly-clicked block/band.
            # We check via click_nonce: if the nonce for active_sel was just bumped
            # (i.e., active_sel was set in a VERY recent prior invocation), skip.
            if active_sel:
                nonce_now = click_nonce_map.get(active_sel, 0)
                nonce_prev = (click_nonce_state or {}).get("nonce", {}).get(active_sel, 0)
                just_set = (nonce_now != nonce_prev)
            else:
                just_set = False
            if not just_set:
                active_sel = None

        # --- C) Selection events: only used to support "re-click same core to cancel highlight" ---
        if trig_selected:
            # if user deselects and the last clickData still points to the same core, cancel it
            if is_deselect(selectedData):
                click_id = resolve_id(clickData)
                if click_id and (click_id.startswith("BLOCK::") or click_id.startswith("BAND::")):
                    if active_sel == click_id:
                        active_sel = None
            # IMPORTANT: do nothing for BTNEXO here (avoid double-toggles)

        # sanitize
        hidden_slices = [s for s in hidden_slices if s in set(slices)]

        # if active selection vanished after rebuild, clear it
        valid_ids = set()
        for oid in curve_to_id:
            base_oid = oid.replace("::TEXT", "") if oid.endswith("::TEXT") else oid
            valid_ids.add(base_oid)
        if active_sel and active_sel not in valid_ids:
            active_sel = None

        # --- Build final figure ---
        fig_out = base_fig

        fig_out, exo_hidden_bands = apply_exo_visibility(fig_out, meta, hidden_slices)

        # keep exo buttons repeat-clickable
        fig_out = apply_btn_nonces(fig_out, meta, btn_nonce_map)

        # keep BG/BLOCK/BAND repeat-clickable (enables "click same again to clear")
        fig_out = apply_click_nonces(fig_out, meta, click_nonce_map)

        # Apply enclosure label visibility (annotations) — this shifts block x-coords
        # for cross-level shift. apply_highlight MUST come after this so that the
        # glow/emboss overlay reads the already-shifted block positions.
        fig_out = apply_enclosure_label_visibility(fig_out, meta, level_state or {"show": {}}, cfg, style_state or {}, graph_size=graph_size)

        # Stamp annotation.name with the original full text on every
        # truncated/wrapped label (consumed by the JS tooltip layer).
        # Same helper is called at initial build time — see _apply_truncation_names.
        _apply_truncation_names(fig_out, meta)

        # Apply highlight AFTER all positional transforms, so glow/emboss overlays
        # are drawn at the correct (post-shift) block coordinates.
        if active_sel:
            fig_out = apply_highlight(fig_out, meta, active_sel, cfg)

        # Apply enclosure border visibility (shapes)
        fig_out = apply_enclosure_border_visibility(fig_out, meta, enclosure_visibility or {})

        # Apply slice label style (annotations)
        fig_out = apply_slice_label_style(fig_out, meta, slice_style_state or {}, rename_map=slice_rename_data or {})

        # Sync collapse line right endpoint to the actual slice label anchor.
        # The label x = enclosure_global_x1 + dist (set by apply_slice_label_style above).
        # build_figure used the default dist; patch the shape here with the real value.
        try:
            _cl_idx = meta.get("collapse_line_shape_idx")
            if _cl_idx is not None:
                _enc_x1 = float(meta.get("enclosure_global_x1") or 0.0)
                _dist   = float((slice_style_state or {}).get("distance", 2.0))
                _cl_x1  = _enc_x1 + _dist
                _shapes = (fig_out.get("layout", {}) or {}).get("shapes") or []
                _ci = int(_cl_idx)
                if 0 <= _ci < len(_shapes):
                    _shapes[_ci]["x1"] = _cl_x1
        except Exception:
            pass

        # Apply band visibility (must be AFTER exo visibility, so it can override)
        # Pass exo_hidden_bands to keep bands connected to hidden Exo blocks hidden
        fig_out = apply_band_visibility(fig_out, meta, {"hidden_types": sorted(list(hidden_types))}, exo_hidden_bands)

        cur_type = str("All")
        # Global button label: show "Show bands" if any hidden, else "Hide bands"
        if len(hidden_types) > 0:
            btn_bands_text = "Show bands"
        else:
            btn_bands_text = "Hide bands"

        # Button label should reflect NEXT action: if all Exo are hidden -> show; else -> hide
        btn_exo_text = "Show Exo" if (len(slices) > 0 and len(hidden_slices) == len(slices)) else "Hide Exo"

        # Only write store-bands when the "Hide/Show all bands" button was the trigger.
        # This breaks the circular Input/Output dependency on store-bands that causes
        # an infinite callback loop on page load.
        # Zoom/pan is preserved purely via JS (restoreZoom in onFigureUpdate MutationObserver).
        # No Python-side axis injection needed — it caused race conditions with the JS restore.
        bands_out = {"hidden_types": sorted(list(hidden_types))} if (
                    trig_btn_bands_all or trig_new_data) else dash.no_update

        # Ensure uirevision is always present.
        if isinstance(fig_out, dict):
            fig_out.setdefault("layout", {})["uirevision"] = "birdcage-fixed"
        else:
            fig_out.update_layout(uirevision="birdcage-fixed")

        # Zoom/pan is now handled entirely client-side via CSS transform on #canvas-content.
        # Python always renders at base zoom (scale=1). No pre-scaling or axis injection needed.

        # ── Interaction log event ──────────────────────────────────────
        import time as _time_ilog
        ilog_evt = dash.no_update
        trig_clickdata_b = any(p.startswith("graph.clickData") for p in trig_set)
        if trig_clickdata_b and clickData:
            _click_id = resolve_id(clickData)
            if _click_id and _click_id.startswith("BLOCK::"):
                evt_type = "block_click"
            elif _click_id and _click_id.startswith("BAND::"):
                evt_type = "band_click"
            else:
                evt_type = None
            if evt_type:
                _params = {}
                if url_search:
                    for _p in url_search.lstrip("?").split("&"):
                        if "=" in _p:
                            _k, _v = _p.split("=", 1)
                            _params[_k] = _v
                _evt_num = 3 if evt_type == "block_click" else 4
                # 用更新后的 active_sel 判断方向：
                # active_sel == _click_id → 刚被设为高亮（highlight）
                # active_sel == None（或其他值）→ 刚被清除（cancel_highlight）
                # 但需排除"第一次点击时 active_sel 恰好与 click_id 相同"的误判：
                # 实际上 active_sel 在 return 之前的值就是 toggle 后的结果，逻辑正确
                _prev_sel = (selected_state or {}).get("selected_id", None)
                if _prev_sel == _click_id:
                    _action = "cancel_highlight"  # 之前已高亮，本次点击取消
                else:
                    _action = "highlight"          # 之前未高亮，本次点击高亮
                ilog_evt = {
                    "type": evt_type,
                    "event_num": _evt_num,
                    "target": _click_id,
                    "action": _action,
                    "ts": int(_time_ilog.time() * 1000),
                    "pid": _params.get("pid", ""),
                    "qid": _params.get("qid", ""),
                    "condition": "A",
                }

        t_render = _t.time() - t0
        print(f"DEBUG: render_with_highlight took {t_render:.3f}s (trig by {trig_main})")
        
        return (
            fig_out,
            {"selected_id": active_sel},
            {"hidden_slices": hidden_slices, "nonce": btn_nonce_map},
            {"nonce": click_nonce_map},
            bands_out,
            btn_exo_text,
            btn_bands_text,
            ilog_evt,
        )

    @app.callback(
        Output("download-panel", "style"),
        Input("btn-download", "n_clicks"),
        Input("btn-download-close", "n_clicks"),
        State("download-panel", "style"),
        State("store-popover-positions", "data"),
        prevent_initial_call=True,
    )
    def toggle_download_panel(n_open, n_close, current_style, popover_positions):
        ctx = dash.callback_context
        trig = ""
        if ctx and ctx.triggered:
            trig = str(ctx.triggered[0].get("prop_id", "")).split(".", 1)[0]

        base = dict(current_style or {})
        base.setdefault("position", "absolute")
        base.setdefault("top", "100%")
        base.setdefault("left", "0px")
        base.setdefault("marginTop", "6px")
        base.setdefault("zIndex", "30000")
        base.setdefault("backgroundColor", "#FFFFFF")
        base.setdefault("border", "2px solid #CCCCCC")
        base.setdefault("borderRadius", "8px")
        base.setdefault("padding", "10px 12px")
        base.setdefault("minWidth", "360px")
        base.setdefault("boxShadow", "0 6px 18px rgba(0,0,0,0.18)")

        # Restore last dragged position if the user has moved this popover.
        _pos = (popover_positions or {}).get("download-panel") or {}
        if _pos.get("left"):
            base["left"] = _pos["left"]
        if _pos.get("top"):
            base["top"] = _pos["top"]
            base["marginTop"] = "0px"

        if trig == "btn-download-close":
            base["display"] = "none"
            return base

        if base.get("display") == "none" or "display" not in base:
            base["display"] = "block"
        else:
            base["display"] = "none"
        return base

    @app.callback(
        Output("ordering-panel", "style"),
        Input("btn-ordering", "n_clicks"),
        Input("btn-ordering-close", "n_clicks"),
        Input("btn-ordering-confirm", "n_clicks"),
        State("ordering-panel", "style"),
        State("store-popover-positions", "data"),
        prevent_initial_call=True,
    )
    def toggle_ordering_panel(n_open, n_close, n_confirm, current_style, popover_positions):
        ctx = dash.callback_context
        trig = ""
        if ctx and ctx.triggered:
            trig = str(ctx.triggered[0].get("prop_id", "")).split(".", 1)[0]

        base = dict(current_style or {})
        base.setdefault("position", "absolute")
        base.setdefault("top", "100%")
        base.setdefault("left", "0px")
        base.setdefault("marginTop", "6px")
        base.setdefault("zIndex", "30000")
        base.setdefault("backgroundColor", "#FFFFFF")
        base.setdefault("border", "2px solid #CCCCCC")
        base.setdefault("borderRadius", "8px")
        base.setdefault("padding", "10px 12px")
        base.setdefault("minWidth", "360px")
        base.setdefault("boxShadow", "0 6px 18px rgba(0,0,0,0.18)")

        # Restore last dragged position if the user has moved this popover.
        _pos = (popover_positions or {}).get("ordering-panel") or {}
        if _pos.get("left"):
            base["left"] = _pos["left"]
        if _pos.get("top"):
            base["top"] = _pos["top"]
            base["marginTop"] = "0px"

        # Close on close button or confirm button click
        if trig in ("btn-ordering-close", "btn-ordering-confirm"):
            base["display"] = "none"
            return base

        if base.get("display") == "none" or "display" not in base:
            base["display"] = "block"
        else:
            base["display"] = "none"
        return base

    @app.callback(
        Output("store-sweep-params", "data"),
        Input("btn-ordering-confirm", "n_clicks"),
        State("sweep-k-max", "value"),
        State("sweep-m", "value"),
        State("sweep-delta", "value"),
        State("store-sweep-params", "data"),
        prevent_initial_call=True,
    )
    def update_sweep_params(n_clicks, k_max, m, delta, current_params):
        if not n_clicks:
            raise PreventUpdate
        # Validate and update parameters
        k_max_use = int(k_max) if k_max and k_max >= 1 else 10
        m_use = int(m) if m and m >= 1 else 2
        delta_use = float(delta) if delta and delta > 0 else 0.01
        return {"k_max": k_max_use, "m": m_use, "delta": delta_use}

    app.clientside_callback(
        """
        function(n_clicks) {
            if (!n_clicks) {
                return [window.dash_clientside.no_update, window.dash_clientside.no_update];
            }
            var host = document.getElementById("graph");
            if (!host) {
                return [{"w": null, "h": null}, {"t": Date.now()}];
            }
            var plots = host.getElementsByClassName("js-plotly-plot");
            if (!plots || plots.length === 0) {
                return [{"w": null, "h": null}, {"t": Date.now()}];
            }
            var gd = plots[0];
            if (!gd || !gd._fullLayout) {
                return [{"w": null, "h": null}, {"t": Date.now()}];
            }
            var w = gd._fullLayout.width;
            var h = gd._fullLayout.height;
            return [{"w": w, "h": h}, {"t": Date.now()}];
        }
        """,
        Output("store-graph-size", "data"),
        Output("store-dl-trigger", "data"),
        Input("btn-download-confirm", "n_clicks"),
        prevent_initial_call=True,
    )

    @app.callback(
        Output("download-image", "data"),
        Input("store-dl-trigger", "data"),
        State("graph", "figure"),
        State("dl-format", "value"),
        State("dl-scale", "value"),
        State("dl-name", "value"),
        State("dl-dpi", "value"),
        State("store-meta", "data"),
        State("store-graph-size", "data"),
        prevent_initial_call=True,
    )
    def download_snapshot(_trig, fig_json, fmt, scale, name, dpi, meta, graph_size):
        if not fig_json:
            raise PreventUpdate

        if fmt == "pdf":
            raise PreventUpdate

        fig = go.Figure(fig_json)

        export_w = None
        export_h = None
        if isinstance(graph_size, dict):
            export_w = graph_size.get("w", None)
            export_h = graph_size.get("h", None)

        if not isinstance(export_w, (int, float)) or export_w <= 0:
            export_w = 1800
        if not isinstance(export_h, (int, float)) or export_h <= 0:
            export_h = 1200

        # Ensure DPI is valid
        try:
            dpi = int(dpi) if dpi else 300
            if dpi < 1:
                dpi = 300
        except Exception:
            dpi = 300

        fig.update_layout(
            autosize=False,
            width=int(export_w),
            height=int(export_h),
            font=dict(family="Arial"),
        )

        # Hide all BTNEXO traces before export (these should not appear in downloaded image)
        # Collapse line is now a layout shape, so it appears naturally in downloads
        # (controlled by line_hidden in collapse_state at build time)
        if meta:
            curve_to_id = meta.get("curve_to_id", [])
            for idx, tr in enumerate(fig.data):
                if idx >= len(curve_to_id):
                    continue
                oid = curve_to_id[idx]
                if oid.startswith("BTNEXO::"):
                    tr.visible = False

            data = fig.to_image(format=fmt, scale=scale, width=int(export_w), height=int(export_h))

            # For PNG and JPEG, set DPI metadata using Pillow
            if fmt in ("png", "jpeg", "jpg"):
                try:
                    from PIL import Image
                    import io

                    # Load image from bytes
                    img = Image.open(io.BytesIO(data))

                    # Create output buffer
                    output = io.BytesIO()

                    # Save with DPI metadata
                    if fmt == "png":
                        # PNG uses pHYs chunk for DPI (pixels per meter)
                        # 1 inch = 0.0254 meters, so pixels_per_meter = dpi / 0.0254
                        img.save(output, format="PNG", dpi=(dpi, dpi))
                    else:
                        # JPEG uses JFIF density
                        img.save(output, format="JPEG", dpi=(dpi, dpi), quality=95)

                    data = output.getvalue()
                except Exception:
                    # If Pillow fails, use original data
                    pass

            return dcc.send_bytes(data, f"{name}.{fmt}")

        app.clientside_callback(
            """
            function(n_clicks, fmt, scale, name) {
                if (!n_clicks) {
                    return "";
                }
                if (fmt !== "pdf") {
                    return "";
                }

                var host = document.getElementById("graph");
                if (!host) {
                    return "";
                }
                var plots = host.getElementsByClassName("js-plotly-plot");
                if (!plots || plots.length === 0) {
                    return "";
                }
                var gd = plots[0];
                if (!gd) {
                    return "";
                }

                // Export a static image first, then print the image in a clean window
                var exportScale = scale || 2;

                return Plotly.toImage(gd, {format: "svg", scale: exportScale}).then(function(dataUrl) {
                    var w = window.open("", "_blank");
                    if (!w) {
                        return "";
                    }

                    // Escape user-supplied `name` before interpolating into the
                    // print-window HTML. Without this, a filename like
                    //   "><script>alert(1)</script>
                    // entered in the File name field would execute in the
                    // newly-opened print window. The same escape is applied
                    // to the <title> and any future attributes.
                    function _escHtml(s) {
                        return String(s == null ? "" : s)
                            .replace(/&/g, "&amp;")
                            .replace(/</g, "&lt;")
                            .replace(/>/g, "&gt;")
                            .replace(/"/g, "&quot;")
                            .replace(/'/g, "&#39;");
                    }
                    var safeName = _escHtml(name || "birdcage_diagram");
                    var title = safeName + ".pdf";
                    w.document.open();
                    w.document.write(
                        "<!doctype html><html><head><meta charset='utf-8'>" +
                        "<title>" + title + "</title>" +
                        "<style>" +
                        "@page{size: landscape; margin: 0;}" +
                        "html,body{margin:0; padding:0;}" +
                        "img{display:block; width:100vw; height:auto;}" +
                        "</style>" +
                        "</head><body>" +
                        "<img id='print-img' src='" + dataUrl + "'/>" +
                        "</body></html>"
                    );
                    w.document.close();

                    var img = w.document.getElementById("print-img");
                    if (img) {
                        img.onload = function() {
                            w.focus();
                            w.print();
                        };
                    } else {
                        w.focus();
                        w.print();
                    }

                    return "";
                }).catch(function(e) {
                    return "";
                });
            }
            """,
            Output("print-dummy", "children"),
            Input("btn-download-confirm", "n_clicks"),
            State("dl-format", "value"),
            State("dl-scale", "value"),
            State("dl-name", "value"),
            prevent_initial_call=True,
        )

    # ── Interaction log: Flask route ──────────────────────────────────
    import os as _os, threading as _threading
    _ilog_lock = _threading.Lock()
    _ilog_script_dir = _os.path.dirname(_os.path.abspath(__file__))

    @app.server.route("/log-event", methods=["POST"])
    def _log_event():
        from flask import request as _req
        try:
            payload = _req.get_json(force=True, silent=True) or {}
            pid = str(payload.get("pid", "unknown"))
            qid = str(payload.get("qid", "unknown"))
            import tempfile as _tmp
            log_dir = _os.path.join(_tmp.gettempdir(), "birdcage_exp", "interaction_logs")
            _os.makedirs(log_dir, exist_ok=True)
            fpath = _os.path.join(log_dir, f"ilog_A_pid{pid}_qid{qid}.json")
            with _ilog_lock:
                # 读取现有结构或初始化
                try:
                    with open(fpath, "r", encoding="utf-8") as _f:
                        record = json.load(_f)
                except Exception:
                    record = {"pid": pid, "qid": qid, "condition": "A", "events": []}
                # 解析 target 字段为可读结构
                evt = {k: v for k, v in payload.items() if k not in ("pid", "qid", "condition")}
                target = evt.get("target", "")
                if isinstance(target, str):
                    if target.startswith("BLOCK::"):
                        parts = target.split("::")
                        if len(parts) >= 3:
                            evt["slice"] = parts[1]
                            evt["category"] = parts[2]
                    elif target.startswith("BAND::"):
                        parts = target.split("::")
                        if len(parts) >= 5:
                            evt["src_slice"] = parts[1]
                            evt["src_category"] = parts[2]
                            evt["dst_slice"] = parts[3]
                            evt["dst_category"] = parts[4]
                record["events"].append(evt)
                with open(fpath, "w", encoding="utf-8") as _f:
                    json.dump(record, _f, ensure_ascii=False, indent=2)
        except Exception:
            pass
        return ("", 204)

    # ── Interaction log: clientside callback → POST to /log-event + track render latency ─────
    app.clientside_callback(
        """
function(evt) {
    if (!evt || !evt.type) return window.dash_clientside.no_update;
    var base = window.location.origin;
    // Record click timestamps for latency tracking
    if (evt.type === 'block_click' || evt.type === 'band_click' ||
        evt.type === 'collapse_click' || evt.type === 'expand_click' ||
        evt.type === 'category_select') {
        window._birdcage_pending_evt = evt;
        window._birdcage_click_ts = evt.ts;
    }
    fetch(base + '/log-event', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify(evt)
    }).catch(function(){});

    // Register category dropdown open listener once
    if (!window._birdcage_cat_listener) {
        window._birdcage_cat_listener = true;
        document.addEventListener('click', function(e) {
            var el = e.target;
            while (el && el !== document) {
                if (el.classList && el.classList.contains('category-combo-btn')) {
                    // Only fire if dropdown is being opened (not for option selection)
                    var params = {};
                    (window.location.search || '').replace(/[?&]([^=&]+)=([^&]*)/g, function(_, k, v) { params[k] = v; });
                    fetch(base + '/log-event', {
                        method: 'POST',
                        headers: {'Content-Type': 'application/json'},
                        body: JSON.stringify({
                            type: 'category_dropdown_open',
                            event_num: 8,
                            ts: Date.now(),
                            pid: params.pid || '',
                            qid: params.qid || '',
                            condition: 'A'
                        })
                    }).catch(function(){});
                    break;
                }
                el = el.parentElement;
            }
        });
    }
    return window.dash_clientside.no_update;
}
""",
        Output("store-ilog-event", "data"),
        Input("store-ilog-event", "data"),
        prevent_initial_call=True,
    )

    # ── Track render latency after figure updates ─────────────────────
    app.clientside_callback(
        """
function(fig) {
    if (!fig) return null;
    setTimeout(function() {
        if (!window._birdcage_click_ts) return;
        var now = Date.now();
        var latency = now - window._birdcage_click_ts;
        var evt = window._birdcage_pending_evt;
        if (!evt) return;
        var base = window.location.origin;
        fetch(base + '/log-event', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({
                type: evt.type + '_render_latency',
                latency_ms: latency,
                ts: now,
                pid: evt.pid,
                qid: evt.qid,
                condition: 'A'
            })
        }).catch(function(){});
        window._birdcage_click_ts = null;
        window._birdcage_pending_evt = null;
    }, 100);
    return null;
}
""",
        Output("__resize_trigger", "data", allow_duplicate=True),
        Input("graph", "figure"),
        prevent_initial_call=True,
    )

    # ── No-data hint: permanently clear once data is loaded ──────────────
    # Sets display:none AND wipes children so the element is truly inert.
    # Triggered by store-upload-slices (fires as soon as files are accepted).
    @app.callback(
        Output("no-data-hint-html", "children"),
        Output("no-data-hint-html", "style"),
        Input("store-upload-slices", "data"),
        prevent_initial_call=False,
    )
    def toggle_no_data_hint(upload_state):
        _HINT_HIDDEN = {"display": "none", "pointerEvents": "none",
                        "position": "absolute", "visibility": "hidden"}
        has_data = bool((upload_state or {}).get("paths"))
        if has_data:
            # Clear children completely + hide — element is inert from here on
            return [], _HINT_HIDDEN
        # No data: restore text and let JS control size/visibility
        _HINT_VISIBLE = {
            "display": "block",
            "position": "absolute",
            "top": "50%", "left": "50%",
            # Same structure as the JS output (see applyCanvasTransform) so
            # the first JS write is a no-op rather than a layout recompute.
            "transform": "translate(-50%, -50%) scale(1)",
            "transformOrigin": "center center",
            "fontSize": "0.75vw",                   # matches bottom-left hint; JS overrides on zoom
            "textAlign": "center",
            "color": "#666666",
            "lineHeight": "1.7",
            "pointerEvents": "none",
            "zIndex": "10",
            "whiteSpace": "nowrap",
        }
        from dash import html as _html
        children = [
            _html.Span("No data loaded.", style={"display": "block"}),
            _html.Span("Click 'Upload files'", style={"display": "block"}),
            _html.Span("to select one or more slice tables.",
                       style={"display": "block"}),
        ]
        return children, _HINT_VISIBLE

    return app


# 9) CLI entry

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--slice", action="append",
                        help="Path to one slice Excel file (or a multi-sheet Excel). Optional when using UI upload.")
    parser.add_argument("--slice-name", action="append",
                        help="Optional display name per slice (same count/order as --slice).")
    parser.add_argument("--category", default=None, help="Default category column name (optional).")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8053)
    parser.add_argument("--debug", action="store_true")
    args = parser.parse_args()

    raw_paths = args.slice or []

    # ── 展开多sheet文件 ──────────────────────────────────────────
    # 如果只传了一个 .xlsx 且它有多个 sheet，自动展开成多个 slice
    expanded_paths: List[str] = []
    expanded_names: List[str] = []

    import tempfile as _tmpfile
    tmp_expand = Path(_tmpfile.gettempdir()) / "birdcage_cli_expand"
    tmp_expand.mkdir(parents=True, exist_ok=True)

    for i, p in enumerate(raw_paths):
        ext = Path(p).suffix.lower()
        if ext in [".xlsx", ".xls", ".ods"]:
            xl = pd.ExcelFile(p)
            sheets = xl.sheet_names
            if len(sheets) > 1:
                # 多sheet：每个sheet写成临时文件
                for sheet in sheets:
                    df = xl.parse(sheet)
                    out = tmp_expand / f"{Path(p).stem}__{sheet}.xlsx"
                    df.to_excel(str(out), index=False)
                    expanded_paths.append(str(out))
                    expanded_names.append(sheet)
            else:
                expanded_paths.append(p)
                name = (args.slice_name[i] if args.slice_name and i < len(args.slice_name)
                        else Path(p).stem)
                expanded_names.append(name)
        else:
            expanded_paths.append(p)
            name = (args.slice_name[i] if args.slice_name and i < len(args.slice_name)
                    else Path(p).stem)
            expanded_names.append(name)

    app = make_app(
        slice_paths=expanded_paths,
        slice_names=expanded_names if expanded_names else args.slice_name,
        element_col="",  # ignored: rightmost column is always used as element column
        default_category_col=args.category,
    )
    app.run(host=args.host, port=args.port, debug=args.debug)


if __name__ == "__main__":
    main()