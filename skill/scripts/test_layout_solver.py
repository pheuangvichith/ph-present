#!/usr/bin/env python3
"""Unit tests for ph-present layout_solver (Thai/Lao segmentation, box-fit, collisions)."""

import unittest
import layout_solver as ls


class TestThaiLaoSegmentation(unittest.TestCase):
    def test_never_splits_after_leading_vowel_or_inside_combining_marks(self):
        sample = "ลดปัญหาตัวอักษรซ้อนทับหรือตกขอบเมื่อเทียบกับภาษาลาว ເມື່ອທຽບກັນ"
        tokens = ls.segment_thai_lao_tokens(sample)
        for tok in tokens:
            # No token should end with an isolated Thai/Lao leading vowel
            if len(tok) == 1:
                self.assertNotIn(tok, ls._ALL_LEADING, f"Isolated leading vowel token: {tok!r}")
            # No token should start with a combining mark
            self.assertFalse(ls.is_combining_thai_lao(tok[0]), f"Token starts with combining mark: {tok!r}")

    def test_chat_bubble_phrase_wraps_cleanly_and_reports_height(self):
        # Exact phrase from the RED-phase baseline test that risked clipping inside h=0.44
        phrase = "มีนมโอ๊ตสูตรไม่หวานส่งด่วนวันนี้ไหมคะ"
        lines = ls.wrap_text_lines(phrase, box_w_in=2.34, size_pt=10.0)
        self.assertGreaterEqual(len(lines), 1)
        for ln in lines:
            self.assertFalse(ls.is_combining_thai_lao(ln[0]))

        res = ls.measure_text_block(
            [ls.TextRunSpec(phrase, size_pt=10.0, line_spacing=1.18)],
            box_w_in=2.58,
            box_h_in=0.44,
            pad_in=(0.12, 0.07, 0.12, 0.07),
        )
        # Ensure solver measures finite positive height
        self.assertGreater(res.required_h_in, 0.25)

    def test_leading_vowel_cluster_and_coda_not_split(self):
        self.assertEqual(ls.segment_thai_lao_tokens("ไทย"), ["ไทย"])
        self.assertEqual(ls.segment_thai_lao_tokens("ใคร"), ["ใคร"])
        self.assertEqual(ls.segment_thai_lao_tokens("โลก"), ["โลก"])
        self.assertEqual(ls.segment_thai_lao_tokens("เพลง"), ["เพลง"])
        self.assertEqual(ls.segment_thai_lao_tokens("ใกล้"), ["ใกล้"])
        self.assertEqual(ls.segment_thai_lao_tokens("ลดค่าใช้จ่าย (บาท)"), ["ลดค่าใช้จ่าย", " ", "(บาท)"])
        self.assertEqual(ls.segment_thai_lao_tokens("ເມືອງລາວ"), ["ເມືອງລາວ"])


class TestConstraintSolver(unittest.TestCase):
    def test_vertical_stack_routes_overflow_without_shrinking_font(self):
        items = [
            ls.StackItem(
                f"card_{i}",
                [
                    ls.TextRunSpec(f"{i}. หัวข้อที่ {i}", size_pt=13.0, bold=True, space_after_pt=2),
                    ls.TextRunSpec(
                        "วิเคราะห์ยอดขายย้อนหลัง ฤดูกาล และโปรโมชันร่วมกับปัจจัยท้องถิ่นเพื่อลดสินค้าค้างสต็อก",
                        size_pt=11.0,
                        line_spacing=1.22,
                    ),
                ],
            )
            for i in range(1, 6)
        ]
        solved = ls.solve_vertical_stack(items, x_in=0.50, w_in=5.00, start_y_in=1.30, max_bottom_in=4.20, gap_in=0.14)
        self.assertGreater(len(solved.placed), 0)
        self.assertGreater(len(solved.overflow_ids), 0)
        for _, y, h in solved.placed:
            self.assertLessEqual(y + h, 4.21)

    def test_aabb_collision_and_whitespace(self):
        boxes = [
            ls.BoxSpec("card_a", 0.50, 1.30, 4.00, 1.50, layer="content"),
            ls.BoxSpec("card_b", 4.20, 1.80, 4.00, 1.50, layer="content"),
        ]
        issues = ls.check_aabb_collisions(boxes)
        self.assertEqual(len(issues), 1)
        self.assertIn("card_a", issues[0])

        clean_boxes = [
            ls.BoxSpec("card_a", 0.50, 1.30, 3.80, 1.50, layer="content"),
            ls.BoxSpec("card_b", 4.60, 1.30, 3.80, 1.50, layer="content"),
        ]
        self.assertEqual(ls.check_aabb_collisions(clean_boxes), [])
        ws = ls.compute_whitespace_ratio(clean_boxes)
        self.assertGreater(ws, 0.75)

        # Unparented content overlapping container must still be caught
        cross_layer = [
            ls.BoxSpec("card_bg", 0.50, 1.30, 3.80, 1.50, layer="container"),
            ls.BoxSpec("rogue_txt", 1.00, 1.50, 2.00, 0.80, layer="content", parent=None),
        ]
        self.assertEqual(len(ls.check_aabb_collisions(cross_layer)), 1)

    def test_content_inventory_rejects_unverified_metrics(self):
        inv = [
            ls.SlideContentInventory(1, "AI Retail 2027", ["3 pillars"], "Overview"),
            ls.SlideContentInventory(2, "Inventory AI", ["-32% waste"], "Cut spoilage", unverified_metrics=[]),
            ls.SlideContentInventory(3, "Thai Chat AI", ["+24% close"], "24/7 commerce", unverified_metrics=["12 boxes"]),
        ]
        errs = ls.validate_content_inventory(inv)
        self.assertEqual(len(errs), 1)
        self.assertIn("unverified synthetic metrics", errs[0])


if __name__ == "__main__":
    unittest.main()
