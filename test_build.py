import json
import tempfile
import unittest
import zipfile
from collections import Counter
from pathlib import Path

import build


class VolubilisYomitanBuildTests(unittest.TestCase):
    def test_split_top_level_alternatives_preserves_bracketed_pronunciation(self) -> None:
        self.assertEqual(build.split_top_level_alternatives("-phē-lā [= -phlao]"), ["-phē-lā [= -phlao]"])
        self.assertEqual(build.split_top_level_alternatives("_a-dīt ; _a-dit"), ["_a-dīt", "_a-dit"])

    def test_split_thai_headwords_requires_thai_in_every_alias(self) -> None:
        self.assertEqual(build.split_thai_headwords("อดีตชาติ ; อดีตภพ"), ["อดีตชาติ", "อดีตภพ"])
        self.assertEqual(build.split_thai_headwords("ไทย = English"), ["ไทย = English"])

    def test_tone_marked_romanization(self) -> None:
        self.assertEqual(build.tone_marked_romanization("-mā"), "maa")
        self.assertEqual(build.tone_marked_romanization("¯mā"), "máa")
        self.assertEqual(build.tone_marked_romanization("/mā"), "mǎa")
        self.assertEqual(build.tone_marked_romanization("_sa_wat-dī"), "sàwàtdii")
        self.assertEqual(build.tone_marked_romanization("-phē-lā [= -phlao]"), "pheelaa [= phlao]")

    def test_pair_headwords_and_readings(self) -> None:
        self.assertEqual(
            build.pair_headwords_and_readings("อดีตชาติ ; อดีตภพ", "_a-dīt ; _a-dit"),
            [
                ("อดีตชาติ", "àdiit", ("อดีตภพ",)),
                ("อดีตภพ", "àdit", ("อดีตชาติ",)),
            ],
        )

    def test_structured_glossary_uses_list_for_multiple_senses(self) -> None:
        definition = build.structured_glossary(
            [
                build.Sense(english="fire", part_of_speech="n."),
                build.Sense(english="electric light", part_of_speech="n."),
            ]
        )
        self.assertEqual(definition["type"], "structured-content")
        self.assertEqual(definition["content"]["tag"], "ol")
        build.validate_structured_content(definition["content"])

    def test_tag_bank_coalesces_case_variants(self) -> None:
        tags = build.make_tag_bank(Counter({"X": 2, "x": 1, "n.": 4}))
        self.assertEqual([row[0] for row in tags], ["pos-n", "pos-x"])
        self.assertEqual(tags[1][3], "X / x")

    def test_deterministic_zip(self) -> None:
        files = {"index.json": b"{}\n", "term_bank_1.json": b"[]\n"}
        with tempfile.TemporaryDirectory() as temporary_directory:
            first = Path(temporary_directory) / "first.zip"
            second = Path(temporary_directory) / "second.zip"
            build.write_deterministic_zip(first, files)
            build.write_deterministic_zip(second, files)
            self.assertEqual(first.read_bytes(), second.read_bytes())
            with zipfile.ZipFile(first) as archive:
                self.assertEqual(json.loads(archive.read("index.json")), {})


if __name__ == "__main__":
    unittest.main()
