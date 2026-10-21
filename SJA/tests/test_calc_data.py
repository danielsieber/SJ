import unittest

from src.pipeline.phase_20_preprocessing.calc_data import create_fol_calcdata


def fol_rows(bibs):
    rows = []
    for index, bib in enumerate(bibs):
        row = [None] * 28
        row[0] = "same-codex"
        row[3] = "2026-01-01"
        row[4] = f"10:00:{index:02d}"
        row[5] = 0
        row[6] = bib
        rows.append(row)
    return rows


def dg_starts(bibs):
    result = create_fol_calcdata(fol_rows(bibs))
    return [index for index, row in enumerate(result) if row[0] is not None]


class CompactTeamBibTests(unittest.TestCase):
    def test_compact_team_suffix_reset_starts_new_dg(self):
        self.assertEqual(dg_starts([11, 21, 31, 124, 11]), [0, 4])

    def test_compact_team_suffix_increase_stays_in_same_dg(self):
        self.assertEqual(dg_starts([11, 21, 31, 42]), [0])

    def test_full_compact_team_event_is_not_split_between_rounds(self):
        bibs = [
            prefix * 10 + suffix
            for suffix in range(1, 5)
            for prefix in range(1, 18)
        ]
        bibs.append(11)

        self.assertEqual(dg_starts(bibs), [0, 68])

    def test_full_explicit_team_event_is_not_split_between_rounds(self):
        bibs = [
            f"{prefix}-{suffix}"
            for suffix in range(1, 5)
            for prefix in range(1, 18)
        ]
        bibs.append("1-1")

        self.assertEqual(dg_starts(bibs), [0, 68])

    def test_compact_and_explicit_team_formats_can_share_one_codex(self):
        compact = [
            prefix * 10 + suffix
            for suffix in range(1, 5)
            for prefix in range(1, 18)
        ]
        explicit = [
            f"{prefix}-{suffix}"
            for suffix in range(1, 5)
            for prefix in range(1, 18)
        ]

        self.assertEqual(dg_starts(compact + explicit), [0, 68])

    def test_normal_bibs_are_not_split_without_compact_team_pattern(self):
        self.assertEqual(dg_starts([10, 11, 12, 13, 4]), [0])

    def test_explicit_team_bibs_keep_existing_suffix_rule(self):
        self.assertEqual(dg_starts(["1-1", "2-1", "3-1", "12-4", "1-1"]), [0, 4])


if __name__ == "__main__":
    unittest.main()
