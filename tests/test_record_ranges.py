import unittest
from dataclasses import replace
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from app import (ChatRecord, OutputOptions, SourceInput, SpeakerSettings,
                 build_docx, build_output_items, collect_records, group_record_ranges, group_time_ranges, parse_range_time)
from docx import Document
from streamlit.testing.v1 import AppTest


class RecordRangeTests(unittest.TestCase):
    def setUp(self):
        self.options = OutputOptions(True, True, False, False, False, False, False, "#808080")
        self.sources = [
            SourceInput("A.txt", "\n".join(f"甲 2026-11-22 10:00:0{i}\nA{i}" for i in (0, 2, 4, 6))),
            SourceInput("B.txt", "\n".join(f"乙 2026-11-22 10:00:0{i}\nB{i}" for i in (1, 3, 5, 7))),
        ]
        self.records = collect_records(self.sources, self.options)

    def contents(self, ranges):
        return [record.content for record in group_record_ranges(self.records, ranges)]

    def test_boundaries_and_multiple_ranges(self):
        self.assertEqual(self.contents([]), [r.content for r in self.records])
        self.assertEqual(self.contents([(1, 4)]), ["A0", "A2", "A4", "B1", "B3", "B5", "A6", "B7"])
        self.assertEqual(self.contents([(4, 7), (0, 3)]), ["A0", "A2", "B1", "B3", "A4", "A6", "B5", "B7"])
        self.assertEqual(self.contents([(2, 2)]), self.contents([]))

    def test_overlap_duplicates_and_invalid_ranges(self):
        self.assertEqual(self.contents([(1, 3), (3, 6), (1, 3)]), self.contents([(1, 6)]))
        for interval in [(-1, 2), (3, 2), (0, 8)]:
            with self.assertRaises(ValueError):
                self.contents([interval])
        self.assertEqual(len(group_record_ranges(self.records, [(0, 7)] * 100)), 8)

    def test_time_ranges_include_all_records_at_boundary(self):
        duplicate = replace(self.records[1], timestamp=self.records[2].timestamp)
        records = sorted([*self.records, duplicate], key=lambda record: record.timestamp)
        result = group_time_ranges(records, [(self.records[2].timestamp, self.records[4].timestamp)])
        self.assertEqual([r.content for r in result], ["A0", "B1", "A2", "A4", "B1", "B3", "B5", "A6", "B7"])
        self.assertEqual(len(result), len(records))
        self.assertEqual(group_time_ranges(records, []), records)

    def test_export_order_filtering_and_headings(self):
        ordered = group_record_ranges(self.records, [(0, 7)])
        settings = {name: SpeakerSettings(name, "#000000", False) for name in ("甲", "乙")}
        options = replace(self.options, add_file_headings=True)
        items = build_output_items(ordered, options, {"A.txt": "A", "B.txt": "B"}, settings)
        doc = Document(build_docx(items, {s.name: "#FFFFFF" for s in self.sources}, settings, options))
        self.assertEqual([p.text for p in doc.paragraphs], ["A", "甲：A0", "甲：A2", "甲：A4", "甲：A6", "B", "乙：B1", "乙：B3", "乙：B5", "乙：B7"])
        settings["甲"].delete = True
        self.assertEqual([i.content for i in build_output_items(ordered, self.options, {}, settings)], ["B1", "B3", "B5", "B7"])

    def test_manual_time_validation_and_optional_minutes_seconds(self):
        self.assertEqual(parse_range_time(["2026", "11", "22", "10", "", ""]), (2026, 11, 22, 10, 0, 0))
        self.assertEqual(parse_range_time(["2024", "2", "29", "0", " 5 ", ""]), (2024, 2, 29, 0, 5, 0))
        self.assertIsNone(parse_range_time(["", "11", "22", "10", "", ""]))
        for values in (["2026", "2", "29", "10", "", ""], ["2026", "11", "22", "24", "", ""], ["2026", "11", "22", "10", "60", ""], ["2026", "11", "22", "10", "", "abc"]):
            with self.assertRaises(ValueError):
                parse_range_time(values)

    def test_ui_add_select_remove_and_export(self):
        at = AppTest.from_file(str(Path(__file__).resolve().parents[1] / "app.py")).run()
        at.radio[0].set_value("粘贴记录").run()
        at.number_input[0].set_value(2).run()
        for index, source in enumerate(self.sources):
            at.text_area[index].set_value(source.raw)
        at.run()
        def button(label):
            return next(b for b in at.button if b.label == label)
        button("添加区间").click().run()
        self.assertTrue(button("生成 Word 文档").disabled)
        def fill_time(key, values):
            for index, value in enumerate(values):
                at.text_input(key=f"{key}-{index}").set_value(value)
            at.run()
        fill_time("range-start-0", ["2026", "11", "22", "10", "", ""])
        fill_time("range-end-0", ["2026", "11", "22", "10", "", "7"])
        self.assertFalse(button("生成 Word 文档").disabled)
        button("生成 Word 文档").click().run()
        self.assertTrue(at.success)
        button("添加区间").click().run()
        self.assertEqual(len([w for w in at.text_input if w.key and w.key.startswith("range-")]), 24)
        at.button(key="remove-range-0").click().run()
        self.assertEqual(len([w for w in at.text_input if w.key and w.key.startswith("range-")]), 12)
        fill_time("range-start-1", ["2026", "11", "22", "10", "", "7"])
        fill_time("range-end-1", ["2026", "11", "22", "10", "", ""])
        self.assertTrue(button("生成 Word 文档").disabled)
        at.text_area[0].set_value(self.sources[0].raw + "\n甲 2026-11-22 10:01:00\n新增").run()
        self.assertEqual(len([w for w in at.text_input if w.key and w.key.startswith("range-")]), 0)
        self.assertFalse(at.exception, at.exception)


if __name__ == "__main__":
    unittest.main()
