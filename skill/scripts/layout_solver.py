#!/usr/bin/env python3
"""Pre-build & post-build Layout, Thai/Lao Wrapping, and Collision Solver for ph-present.

Zero external dependencies beyond Python stdlib + Pillow (optional; falls back to
conservative glyph-width estimation if Pillow is absent).
"""

from __future__ import annotations

import os
import re
from dataclasses import dataclass, field

try:
    from PIL import ImageFont
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

EMU_PER_INCH = 914400.0
SLIDE_W_IN = 10.0
SLIDE_H_IN = 5.625

# Thai/Lao Unicode character classes
_THAI_LEADING = set("เแโไใ")
_LAO_LEADING = set("ເແໂໄໃ")
_ALL_LEADING = _THAI_LEADING | _LAO_LEADING

_THAI_FOLLOW = set("ะาำๅๆฯ")
_LAO_FOLLOW = set("ະາຳໆຯ")
_ALL_FOLLOW = _THAI_FOLLOW | _LAO_FOLLOW

_NO_BREAK_BEFORE = set(".,;:!?%)]}\"'”’ๆฯໆຯ์")
_NO_BREAK_AFTER = set("([{\"'“‘")

FONT_CANDIDATES = {
    ("latin", False): [
        "/usr/share/fonts/liberation/LiberationSans-Regular.ttf",
        "/usr/share/fonts/TTF/DejaVuSans.ttf",
    ],
    ("latin", True): [
        "/usr/share/fonts/liberation/LiberationSans-Bold.ttf",
        "/usr/share/fonts/TTF/DejaVuSans-Bold.ttf",
    ],
    ("thai", False): [
        "/usr/share/fonts/noto/NotoSansThai-Regular.ttf",
        "/usr/share/fonts/truetype/noto/NotoSansThai-Regular.ttf",
    ],
    ("thai", True): [
        "/usr/share/fonts/noto/NotoSansThai-Bold.ttf",
        "/usr/share/fonts/truetype/noto/NotoSansThai-Bold.ttf",
    ],
    ("lao", False): [
        "/usr/share/fonts/noto/NotoSansLao-Regular.ttf",
        "/usr/share/fonts/truetype/noto/NotoSansLao-Regular.ttf",
    ],
    ("lao", True): [
        "/usr/share/fonts/noto/NotoSansLao-Bold.ttf",
        "/usr/share/fonts/truetype/noto/NotoSansLao-Bold.ttf",
    ],
    ("mono", False): [
        "/usr/share/fonts/TTF/DejaVuSansMono.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf",
    ],
    ("mono", True): [
        "/usr/share/fonts/TTF/DejaVuSansMono-Bold.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSansMono-Bold.ttf",
    ],
}

_FONT_CACHE: dict[tuple[str, bool, int], object] = {}

# High-frequency Thai/Lao lexicon for DP word segmentation (paired with orthographic syllable fallback)
_COMMON_WORDS = {
    # Core connectors, prepositions, modifiers, short function words
    "และ", "หรือ", "เพื่อ", "กับ", "จาก", "การ", "ความ", "เป็น", "ได้", "ให้", "ใน", "บน", "ล่าง",
    "ของ", "ที่", "ซึ่ง", "โดย", "ตาม", "ด้วย", "ผ่าน", "ข้าม", "ถึง", "ก่อน", "หลัง", "ระหว่าง",
    "มี", "คือ", "นี้", "นั้น", "ไป", "มา", "ดู", "ส่ง", "ซื้อ", "ขาย", "จ่าย", "คิด", "รับ", "เพิ่ม",
    "ทุก", "ทั้ง", "ทั้งหมด", "แต่ละ", "ไม่", "ไม่มี", "ยัง", "กำลัง", "แล้ว", "เมื่อ", "ถ้า", "หาก",
    "ร่วม", "ช่วย", "ทำ", "ใช้", "สร้าง", "แก้", "ลด", "เพิ่ม", "ปรับ", "ตรวจ", "เทียบ", "แบ่ง", "วาง",
    "จริง", "แม่นยำ", "รวดเร็ว", "อัตโนมัติ", "อัจฉริยะ", "ต่อเนื่อง", "ธรรมชาติ", "สมบูรณ์", "ชัดเจน",
    "มาก", "น้อย", "สูง", "ต่ำ", "ใหญ่", "เล็ก", "ใหม่", "เดิม", "แรก", "เดียว", "ครบ", "พร้อม",
    "แบบ", "ระดับ", "ด้าน", "หลัก", "ย่อย", "เฉพาะ", "ทั่วไป", "ภายใน", "ภายนอก", "ตรง", "กัน",
    # Business, retail, tech, presentation vocabulary
    "กลยุทธ์", "ธุรกิจ", "ค้าปลีก", "ระบบ", "คาดการณ์", "สต็อก", "ของเสีย", "ลดของเสีย", "สินค้า",
    "คลังสินค้า", "ผู้ช่วย", "ภาษาไทย", "ภาษาลาว", "ปิดการขาย", "แดชบอร์ด", "สาขา", "ประสิทธิภาพ",
    "ผลลัพธ์", "ข้อมูล", "ลูกค้า", "โปรโมชัน", "เรียลไทม์", "มาตรฐาน", "ตรวจสอบ", "ข้อผิดพลาด",
    "สไลด์", "ตัวอักษร", "ภาพรวม", "โครงสร้าง", "สถาปัตยกรรม", "เครื่องมือ", "องค์กร", "ประเทศไทย",
    "การตลาด", "ยอดขาย", "วิเคราะห์", "พยากรณ์", "อาหารสด", "ชำระเงิน", "แนะนำ", "ตัวอย่าง",
    "เปรียบเทียบ", "ข้อความ", "ซ้อนทับ", "ตกขอบ", "โปรเจกเตอร์", "คุณภาพ", "ปัญหา", "สาเหตุ",
    "ปัจจัย", "ท้องถิ่น", "ย้อนหลัง", "ฤดูกาล", "ค้าง", "ค้างสต็อก", "หมดอายุ", "วันหยุด", "เทศกาล",
    "เงินเดือน", "ความต้องการ", "อุปสงค์", "คำถาม", "คำตอบ", "บทสนทนา", "ออเดอร์", "แอดมิน",
    "ทีมขาย", "โปร่งใส", "ไร้รอยต่อ", "คิดเงิน", "หน้าร้าน", "ห่วงโซ่", "อุปทาน", "ปฏิบัติการ",
    "เนื้อหา", "ประเด็น", "สำคัญ", "สรุป", "ออกแบบ", "หน้าตา", "จังหวะ", "การเคลื่อนไหว", "หมุน",
    "ซ้าย", "ขวา", "กลาง", "บน", "ล่าง", "ขอบ", "กรอบ", "การ์ด", "กราฟ", "แผนภาพ", "รูปภาพ",
    "วันนี้", "พรุ่งนี้", "ส่งด่วน", "ส่งฟรี", "สูตร", "หวาน", "ไม่หวาน", "นมโอ๊ต", "แพ็กคู่",
    "กล่อง", "ชิ้น", "หน้า", "รอบ", "เวลา", "ชั่วโมง", "นาที", "วินาที", "ไหมคะ", "นะคะ", "ครับ", "ค่ะ",
    "โปร่ง", "สะอาด", "ทันสมัย", "นำเสนอ", "ตัดสินใจ", "ผู้บริหาร", "เป้าหมาย", "ตัวชี้วัด",
    "ไทย", "ใคร", "โลก", "เพลง", "ใกล้", "งาน", "บ้าน", "บาท", "ไตรมาส", "รายได้", "กำไร",
    "ต้นทุน", "เติบโต", "พัฒนา", "สรุป", "ค่าใช้จ่าย", "ลดค่าใช้จ่าย", "รายเดือน", "ประจำ",
    # Lao high-frequency words
    "ສາຂາ", "ລະບົບ", "ຂໍ້ມູນ", "ລູກຄ້າ", "ທຸລະກິດ", "ພາສາລາວ", "ກວດສອບ", "ທົດສອບ",
    "ປະສິດທິພາບ", "ຜົນລັບ", "ສະຫຼຸບ", "ໂຄງສ້າງ", "ຮູບພາບ", "ຂໍ້ຄວາມ", "ພ້ອມ", "ແລະ", "ຫຼື",
    "ເພື່ອ", "ກັບ", "ຈາກ", "ການ", "ຄວາມ", "ເປັນ", "ໄດ້", "ໃນ", "ຂອງ", "ທີ່", "ເມື່ອ", "ທຽບກັນ",
    "ສະບາຍດີ", "ຂອບໃຈ", "ເມືອງລາວ", "ພັດທະນາ", "ລາຍໄດ້",
}
_MAX_WORD_LEN = max(len(w) for w in _COMMON_WORDS)


def is_thai(ch: str) -> bool:
    return "\u0E00" <= ch <= "\u0E7F"


def is_lao(ch: str) -> bool:
    return "\u0E80" <= ch <= "\u0EFF"


def is_combining_thai_lao(ch: str) -> bool:
    c = ord(ch)
    return (
        c in (0x0E31, 0x0EB1)
        or (0x0E34 <= c <= 0x0E3A)
        or (0x0E47 <= c <= 0x0E4E)
        or (0x0EB4 <= c <= 0x0EBC)
        or (0x0EC8 <= c <= 0x0ECD)
    )


def _is_thai_lao_consonant(ch: str) -> bool:
    c = ord(ch)
    return (0x0E01 <= c <= 0x0E2E) or (0x0E81 <= c <= 0x0EAE)


def _resolve_font_path(kind: str, bold: bool) -> str | None:
    for p in FONT_CANDIDATES.get((kind, bold), []):
        if os.path.exists(p):
            return p
    for p in FONT_CANDIDATES.get((kind, False), []):
        if os.path.exists(p):
            return p
    for p in FONT_CANDIDATES[("latin", bold)]:
        if os.path.exists(p):
            return p
    return None


def get_ttf_font(kind: str, bold: bool, pt: float, dpi: float = 96.0):
    if not HAS_PIL:
        return None
    px = max(8, int(round(pt * (dpi / 72.0))))
    key = (kind, bold, px)
    if key not in _FONT_CACHE:
        path = _resolve_font_path(kind, bold)
        _FONT_CACHE[key] = ImageFont.truetype(path, px) if path else ImageFont.load_default()
    return _FONT_CACHE[key]


def measure_token_inches(text: str, kind: str, bold: bool, pt: float) -> float:
    """Measure exact rendered width in inches using installed TTF metrics."""
    if not text:
        return 0.0
    dpi = 128.0
    font = get_ttf_font(kind, bold, pt, dpi=dpi)
    if font is not None and hasattr(font, "getlength"):
        return float(font.getlength(text)) / dpi
    # Fallback heuristic if Pillow/TTF unavailable
    avg_em = 0.54 if kind != "mono" else 0.60
    return len(text) * (pt / 72.0) * avg_em


def split_clusters(text: str) -> list[str]:
    """Group base characters with any Thai/Lao combining vowels/tones/marks."""
    clusters: list[str] = []
    for ch in text:
        if clusters and is_combining_thai_lao(ch):
            clusters[-1] += ch
        else:
            clusters.append(ch)
    return clusters


_CLUSTER_FIRSTS = set("กขคตปผพກຂຄຕປຜພ")
_CLUSTER_SECONDS = set("รลวຣລວ")


def _has_combining_vowel(text: str) -> bool:
    for ch in text:
        c = ord(ch)
        if c in (0x0E31, 0x0EB1) or (0x0E34 <= c <= 0x0E3A) or (0x0EB4 <= c <= 0x0EBC) or c == 0x0E47:
            return True
    return False


def _orthographic_syllables(run: str) -> list[str]:
    """Segment a Thai/Lao run into syllable-like units without breaking vowels/coda."""
    clusters = split_clusters(run)
    raw_units: list[str] = []
    c_idx = 0
    while c_idx < len(clusters):
        unit = clusters[c_idx]
        c_idx += 1
        if unit[0] in _ALL_LEADING and c_idx < len(clusters):
            unit += clusters[c_idx]
            c_idx += 1
            # Attach true onset cluster second consonant (e.g., ใกล้, เพลง, ไตร)
            if (
                c_idx < len(clusters)
                and len(unit) >= 2
                and unit[1] in _CLUSTER_FIRSTS
                and clusters[c_idx][0] in _CLUSTER_SECONDS
                and not _has_combining_vowel(clusters[c_idx])
                and (
                    any(0x0E48 <= ord(ch) <= 0x0E4B or 0x0EC8 <= ord(ch) <= 0x0ECB for ch in clusters[c_idx])
                    or (c_idx + 1 < len(clusters) and _is_thai_lao_consonant(clusters[c_idx + 1][0]) and len(clusters[c_idx + 1]) == 1)
                )
            ):
                unit += clusters[c_idx]
                c_idx += 1
        while c_idx < len(clusters):
            nxt = clusters[c_idx]
            if nxt[0] in _ALL_FOLLOW or "์" in nxt or nxt in ("ๆ", "ฯ", "ໆ", "ຯ"):
                unit += nxt
                c_idx += 1
            else:
                break
        raw_units.append(unit)

    # Glue single bare coda consonant onto an open/vowel-marked unit without an existing coda
    glued: list[str] = []
    for idx, u in enumerate(raw_units):
        is_bare_consonant = len(u) == 1 and _is_thai_lao_consonant(u)
        if is_bare_consonant and glued:
            prev = glued[-1]
            prev_clusters = split_clusters(prev)
            has_vowel = any(ch in _ALL_LEADING or ch in _ALL_FOLLOW for ch in prev) or _has_combining_vowel(prev)
            # Check that prev doesn't already end with a bare coda after a vowel
            ends_with_vowel_or_single_onset = (
                prev[-1] in _ALL_FOLLOW
                or is_combining_thai_lao(prev[-1])
                or (prev[0] in _ALL_LEADING and len(prev_clusters) == 2)
            )
            # Don't steal an onset if the next unit starts with a combining mark? (raw_units already grouped clusters)
            # Also avoid gluing if prev is a closed leading-vowel+cluster onset unless it still lacks a coda
            is_leading_cluster_without_coda = (
                prev[0] in _ALL_LEADING
                and len(prev_clusters) == 3
                and prev_clusters[2][0] in _CLUSTER_SECONDS
                and not prev[-1] in _ALL_FOLLOW
            )
            if has_vowel and (ends_with_vowel_or_single_onset or is_leading_cluster_without_coda):
                # Do not glue if prev is a bare leading+C followed by something like ไท|ย where ไท has no follow vowel -> ok to glue!
                # Wait: prevent over-gluing two distinct syllables only when u is clearly followed by another coda? No, raw_units already attached follow/combining marks to u if u had its own vowel!
                glued[-1] += u
                continue
        glued.append(u)
    return glued


def _dp_segment_run(run: str) -> list[str]:
    """Dynamic-programming word + orthographic syllable segmentation for Thai/Lao runs."""
    n = len(run)
    if n == 0:
        return []
    # Precompute syllable units so unknown spans fall back cleanly to syllables
    syls = _orthographic_syllables(run)
    syl_starts: dict[int, list[str]] = {}
    pos = 0
    for s in syls:
        syl_starts.setdefault(pos, []).append(s)
        pos += len(s)

    # dp[i] = (min_cost, next_index) from index i to n
    INF = 10 ** 9
    dp: list[tuple[int, int]] = [(INF, n)] * (n + 1)
    dp[n] = (0, n)

    for i in range(n - 1, -1, -1):
        best_cost = INF
        best_nxt = i + 1
        # Dictionary matches get very low cost (prefer longer known words)
        max_l = min(_MAX_WORD_LEN, n - i)
        for length in range(max_l, 1, -1):
            sub = run[i : i + length]
            if sub in _COMMON_WORDS:
                # Never leave an orphan combining mark after the match
                if i + length < n and is_combining_thai_lao(run[i + length]):
                    continue
                cost = 1 + dp[i + length][0]
                if cost < best_cost:
                    best_cost = cost
                    best_nxt = i + length
        # Orthographic syllable fallback
        for s in syl_starts.get(i, ()):
            nxt = i + len(s)
            cost = 5 + dp[nxt][0]
            if cost < best_cost:
                best_cost = cost
                best_nxt = nxt
        if best_cost >= INF:
            # Should rarely trigger (only inside a syllable if dict matched awkwardly)
            nxt = i + 1
            while nxt < n and is_combining_thai_lao(run[nxt]):
                nxt += 1
            best_cost = 20 + dp[nxt][0]
            best_nxt = nxt
        dp[i] = (best_cost, best_nxt)

    out: list[str] = []
    cur = 0
    while cur < n:
        nxt = dp[cur][1]
        out.append(run[cur:nxt])
        cur = nxt
    return out


def segment_thai_lao_tokens(text: str) -> list[str]:
    """Split text into wrap-safe tokens without breaking Thai/Lao syllables or leading vowels."""
    if not text:
        return []

    tokens: list[str] = []
    n = len(text)
    i = 0

    while i < n:
        ch = text[i]
        if ch == "\n":
            tokens.append("\n")
            i += 1
            continue
        if ch == " ":
            j = i
            while j < n and text[j] == " ":
                j += 1
            tokens.append(text[i:j])
            i = j
            continue

        if not (is_thai(ch) or is_lao(ch)):
            j = i + 1
            while j < n and text[j] not in (" ", "\n") and not (is_thai(text[j]) or is_lao(text[j])):
                j += 1
            tokens.append(text[i:j])
            i = j
            continue

        j = i
        while j < n and (is_thai(text[j]) or is_lao(text[j])):
            j += 1
        tokens.extend(_dp_segment_run(text[i:j]))
        i = j

    merged: list[str] = []
    carry_prefix = ""
    for tok in tokens:
        if carry_prefix:
            if tok not in (" ", "\n"):
                tok = carry_prefix + tok
                carry_prefix = ""
            else:
                merged.append(carry_prefix)
                carry_prefix = ""
        if tok not in (" ", "\n") and all(c in _NO_BREAK_AFTER for c in tok):
            carry_prefix = tok
            continue
        if (
            merged
            and tok not in (" ", "\n")
            and tok[0] in _NO_BREAK_BEFORE
            and not merged[-1].endswith("\n")
            and merged[-1].strip() != ""
        ):
            merged[-1] += tok
        else:
            merged.append(tok)
    if carry_prefix:
        merged.append(carry_prefix)
    return merged


def _detect_token_kind(tok: str, mono: bool = False) -> str:
    if mono:
        return "mono"
    if any(is_lao(c) for c in tok):
        return "lao"
    if any(is_thai(c) for c in tok):
        return "thai"
    return "latin"


def wrap_text_lines(
    text: str,
    box_w_in: float,
    size_pt: float,
    bold: bool = False,
    mono: bool = False,
) -> list[str]:
    """Wrap text into visual lines that fit inside `box_w_in` inches."""
    if text == "":
        return [""]
    lines: list[str] = [""]
    cur_w = 0.0
    tokens = segment_thai_lao_tokens(text)

    for tok in tokens:
        if tok == "\n":
            lines.append("")
            cur_w = 0.0
            continue
        kind = _detect_token_kind(tok, mono=mono)
        tw = measure_token_inches(tok, kind, bold, size_pt)
        if tok.strip() == "":
            if lines[-1] == "":
                continue
            if cur_w + tw <= box_w_in:
                lines[-1] += tok
                cur_w += tw
            continue

        if cur_w + tw > box_w_in and lines[-1] != "":
            lines.append(tok)
            cur_w = tw
        elif cur_w + tw > box_w_in and lines[-1] == "":
            # Single token wider than box_w_in -> split by cluster only as last resort
            for cl in split_clusters(tok):
                cw = measure_token_inches(cl, kind, bold, size_pt)
                if cur_w + cw > box_w_in and lines[-1] != "":
                    lines.append(cl)
                    cur_w = cw
                else:
                    lines[-1] += cl
                    cur_w += cw
        else:
            lines[-1] += tok
            cur_w += tw

    return [ln.rstrip() for ln in lines]


@dataclass
class TextRunSpec:
    text: str
    size_pt: float
    bold: bool = False
    mono: bool = False
    line_spacing: float = 1.22
    space_after_pt: float = 0.0


@dataclass
class MeasuredTextResult:
    required_h_in: float
    lines_per_para: list[list[str]]
    total_lines: int
    fits: bool
    overflow_in: float


def measure_text_block(
    paragraphs: list[TextRunSpec],
    box_w_in: float,
    box_h_in: float | None = None,
    pad_in: tuple[float, float, float, float] = (0.14, 0.10, 0.14, 0.10),
    safety_in: float = 0.04,
) -> MeasuredTextResult:
    """Measure exact vertical height required by paragraphs inside `(box_w_in, box_h_in)`."""
    l_pad, t_pad, r_pad, b_pad = pad_in
    inner_w = max(0.25, box_w_in - l_pad - r_pad)
    lines_per_para: list[list[str]] = []
    content_h = 0.0
    total_lines = 0

    for p in paragraphs:
        wrapped = wrap_text_lines(p.text, inner_w, p.size_pt, bold=p.bold, mono=p.mono)
        lines_per_para.append(wrapped)
        n_lines = max(1, len(wrapped))
        total_lines += n_lines
        content_h += n_lines * (p.size_pt / 72.0) * p.line_spacing
        content_h += p.space_after_pt / 72.0

    req_h = t_pad + b_pad + content_h + safety_in
    fits = True if box_h_in is None else (req_h <= box_h_in + 1e-6)
    overflow = 0.0 if (box_h_in is None or fits) else (req_h - box_h_in)
    return MeasuredTextResult(
        required_h_in=round(req_h, 4),
        lines_per_para=lines_per_para,
        total_lines=total_lines,
        fits=fits,
        overflow_in=round(overflow, 4),
    )


@dataclass
class StackItem:
    item_id: str
    paragraphs: list[TextRunSpec]
    min_h_in: float = 0.55
    pad_in: tuple[float, float, float, float] = (0.14, 0.08, 0.14, 0.08)


@dataclass
class SolvedStackResult:
    placed: list[tuple[str, float, float]]  # (item_id, y_in, h_in)
    overflow_ids: list[str]
    used_bottom_in: float


def solve_vertical_stack(
    items: list[StackItem],
    x_in: float,
    w_in: float,
    start_y_in: float,
    max_bottom_in: float,
    gap_in: float = 0.12,
) -> SolvedStackResult:
    """Measure every card in a vertical stack and lay out without overlap or clipping.

    Items that cannot fit above `max_bottom_in` are returned in `overflow_ids`
    so the caller can route them to a Morph continuation sub-slide instead of
    shrinking fonts below the projector floor or deleting content.
    """
    placed: list[tuple[str, float, float]] = []
    overflow_ids: list[str] = []
    cur_y = start_y_in

    for idx, item in enumerate(items):
        m = measure_text_block(item.paragraphs, w_in, None, pad_in=item.pad_in)
        h = max(item.min_h_in, m.required_h_in)
        if cur_y + h <= max_bottom_in + 1e-6 and not overflow_ids:
            placed.append((item.item_id, round(cur_y, 4), round(h, 4)))
            cur_y += h + gap_in
        else:
            overflow_ids.append(item.item_id)

    used_bottom = placed[-1][1] + placed[-1][2] if placed else start_y_in
    return SolvedStackResult(placed=placed, overflow_ids=overflow_ids, used_bottom_in=round(used_bottom, 4))


@dataclass
class BoxSpec:
    name: str
    x: float
    y: float
    w: float
    h: float
    layer: str = "content"  # "bg" | "container" | "content" | "chrome"
    parent: str | None = None


def check_aabb_collisions(boxes: list[BoxSpec], min_gap_in: float = 0.04) -> list[str]:
    """Check boxes on content/container layers for unintended overlap (IoU > 0)."""
    issues: list[str] = []
    visible = [
        b for b in boxes
        if b.layer in ("content", "container")
        and (b.x + b.w > 0 and b.x < SLIDE_W_IN and b.y + b.h > 0 and b.y < SLIDE_H_IN)
    ]

    for i in range(len(visible)):
        a = visible[i]
        for j in range(i + 1, len(visible)):
            b = visible[j]
            if a.parent == b.name or b.parent == a.name:
                continue
            if a.layer != b.layer and a.parent is not None and b.parent is not None and a.parent != b.parent:
                continue
            ox = min(a.x + a.w, b.x + b.w) - max(a.x, b.x)
            oy = min(a.y + a.h, b.y + b.h) - max(a.y, b.y)
            if ox > 0.01 and oy > 0.01:
                issues.append(
                    f"Overlap between '{a.name}' and '{b.name}' ({ox:.2f}x{oy:.2f} in)"
                )
    return issues


def compute_whitespace_ratio(boxes: list[BoxSpec]) -> float:
    """Approximate uncovered canvas fraction (0.0..1.0) from top-level visible cards."""
    grid_w, grid_h = 100, 56
    occupied = 0
    top_boxes = [
        b for b in boxes
        if b.layer in ("container", "content") and b.parent is None
        and b.x + b.w > 0 and b.x < SLIDE_W_IN and b.y + b.h > 0 and b.y < SLIDE_H_IN
    ]
    for gx in range(grid_w):
        px = (gx + 0.5) * (SLIDE_W_IN / grid_w)
        for gy in range(grid_h):
            py = (gy + 0.5) * (SLIDE_H_IN / grid_h)
            if any(b.x <= px <= b.x + b.w and b.y <= py <= b.y + b.h for b in top_boxes):
                occupied += 1
    return round(1.0 - (occupied / (grid_w * grid_h)), 3)


@dataclass
class SlideContentInventory:
    slide_index: int
    core_thesis: str
    proof_points: list[str]
    takeaway: str
    unverified_metrics: list[str] = field(default_factory=list)


def validate_content_inventory(inventory: list[SlideContentInventory]) -> list[str]:
    """Enforce zero-content-loss structural contract before layout compilation."""
    errors: list[str] = []
    if not inventory:
        return ["Content inventory is empty."]
    for s in inventory:
        if not s.core_thesis.strip():
            errors.append(f"Slide {s.slide_index}: missing core_thesis")
        if not s.proof_points and s.slide_index > 1:
            errors.append(f"Slide {s.slide_index}: missing proof_points/evidence")
        if not s.takeaway.strip():
            errors.append(f"Slide {s.slide_index}: missing takeaway")
        if s.unverified_metrics:
            errors.append(f"Slide {s.slide_index}: unverified synthetic metrics flagged: {s.unverified_metrics}")
    return errors
