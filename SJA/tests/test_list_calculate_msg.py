import unittest
from datetime import date, time

from src.pipeline.phase_20_preprocessing.format_save import _build_jumper_id
from src.pipeline.phase_20_preprocessing.list_calculate import (
    _loop4_sync_msg_with_fol,
    find_trialjumper_remark_rows,
    run_calculation,
)


def fol_row(
    dg,
    typ,
    bib,
    codex,
    hour,
    minute,
    *,
    name=None,
    vorname=None,
    nsa=None,
    message=None,
    gap=None,
):
    row = [None] * 33
    row[0] = dg
    row[1] = typ
    row[2] = bib
    row[3] = name
    row[4] = vorname
    row[5] = nsa
    row[26] = message
    row[27] = codex
    row[28] = date(2026, 3, 15)
    row[29] = time(hour, minute)
    row[30] = gap
    return row


def msg_row(
    hour,
    minute,
    message,
    *,
    bib=None,
    status=None,
    gate=None,
    codex=3204,
):
    row = [None] * 10
    row[0] = date(2026, 3, 15)
    row[1] = time(hour, minute)
    row[2] = message
    row[4] = status
    row[5] = bib
    row[6] = gate
    row[7] = codex
    return row


class MessageToFolCalculationTests(unittest.TestCase):
    def setUp(self):
        self.data = [
            fol_row(
                1,
                "W",
                1,
                3204,
                10,
                0,
                name="Alpha",
                vorname="Anna",
                nsa="SUI",
            ),
            fol_row(
                None,
                None,
                2,
                "3204.0",
                10,
                10,
                name="Beta",
                vorname="Ben",
                nsa="AUT",
            ),
            fol_row(
                2,
                "Q",
                1,
                3204,
                10,
                20,
                name="Alpha",
                vorname="Anna",
                nsa="SUI",
            ),
            fol_row(
                None,
                None,
                3,
                3204,
                10,
                30,
                name="Gamma",
                vorname="Gus",
                nsa="GER",
            ),
            fol_row(3, "T", 9, 9999, 10, 4, name="Other"),
            fol_row(None, None, 10, 9999, 10, 6, name="Else"),
        ]

    def sync(self, messages):
        changed = set()
        returned = _loop4_sync_msg_with_fol(self.data, messages, changed)
        return returned

    def test_status_updates_existing_bib_in_same_jump_id(self):
        changed = self.sync(
            [msg_row(10, 21, "Alpha disqualified", bib=1, status="dsq")]
        )

        self.assertIsNone(self.data[0][24])
        self.assertEqual(self.data[2][24], "DSQ")
        self.assertEqual(self.data[2][26], "Alpha disqualified")
        self.assertIn((2, 24), changed)
        self.assertIn((2, 26), changed)

    def test_missing_status_bib_is_inserted_and_enriched_from_other_block(self):
        self.sync(
            [msg_row(10, 25, "Ben Beta did not start", bib=2, status="DNS")]
        )

        inserted = self.data[3]
        self.assertEqual(inserted[2], 2)
        self.assertEqual(inserted[3:6], ["Beta", "Ben", "AUT"])
        self.assertEqual(inserted[24], "DNS")
        self.assertEqual(inserted[26], "Ben Beta did not start")
        self.assertEqual(inserted[27], 3204)
        self.assertEqual(inserted[28], date(2026, 3, 15))
        self.assertEqual(inserted[29], time(10, 25))
        self.assertEqual(
            _build_jumper_id(inserted[3], inserted[4], inserted[5]),
            "AUT BETA BEN",
        )

    def test_identity_stays_empty_when_name_is_not_in_message(self):
        self.sync(
            [msg_row(10, 25, "Athlete did not start", bib=2, status="DNS")]
        )

        self.assertEqual(self.data[3][2], 2)
        self.assertEqual(self.data[3][3:6], [None, None, None])
        self.assertIsNone(_build_jumper_id(*self.data[3][3:6]))

    def test_insertion_before_block_transfers_dg_and_typ_marker(self):
        self.sync(
            [msg_row(10, 16, "Ben Beta not permitted", bib=2, status="NPS")]
        )

        inserted = self.data[2]
        previous_first = self.data[3]
        self.assertEqual(inserted[0:2], [2, "Q"])
        self.assertEqual(inserted[2], 2)
        self.assertEqual(inserted[24], "NPS")
        self.assertEqual(previous_first[0:2], [None, None])

    def test_insertion_after_block_stays_before_next_jump_id(self):
        self.sync(
            [msg_row(10, 40, "Ben Beta did not start", bib=2, status="DNS")]
        )

        self.assertEqual(self.data[4][2], 2)
        self.assertEqual(self.data[4][29], time(10, 40))
        self.assertEqual(self.data[5][0:2], [3, "T"])

    def test_status_without_bib_cannot_insert_a_row(self):
        self.sync(
            [msg_row(10, 25, "Unknown athlete DNS", status="DNS")]
        )

        self.assertEqual(len(self.data), 6)

    def test_bibless_message_uses_nearest_row_and_newest_wins(self):
        self.sync(
            [
                msg_row(10, 4, "newest conditions"),
                msg_row(10, 2, "older conditions"),
            ]
        )

        self.assertEqual(self.data[0][26], "newest conditions")

    def test_control_messages_are_skipped(self):
        for position, text in enumerate(
            (
                "Competition in progress",
                "Start delayed",
                "Trial starts now",
                "Round completed",
            )
        ):
            with self.subTest(text=text):
                data = [fol_row(1, "W", 1, 3204, 10, position)]
                _loop4_sync_msg_with_fol(
                    data,
                    [msg_row(10, position, text)],
                    set(),
                )
                self.assertIsNone(data[0][26])

    def test_trialjumper_moves_from_time_match_to_nearest_large_gap(self):
        data = [
            fol_row(1 if minute == 0 else None, "W", minute + 1, 3204, 10, minute,
                    gap=732 if minute == 5 else 37)
            for minute in range(10)
        ]
        messages = [
            msg_row(10, 4, "Trialjumper", status="trialjumper")
        ]

        changed = set()
        _loop4_sync_msg_with_fol(data, messages, changed)

        self.assertEqual(find_trialjumper_remark_rows(data, messages), {5})
        self.assertTrue(all(row[26] is None for row in data))
        self.assertEqual(changed, set())

    def test_trialjumper_search_includes_four_rows_but_not_five(self):
        data = [
            fol_row(1 if minute == 0 else None, "W", minute + 1, 3204, 10, minute,
                    gap=101 if minute == 0 else (500 if minute == 9 else 20))
            for minute in range(10)
        ]
        messages = [
            msg_row(10, 4, "Trialjumper", status="TRIALJUMPER")
        ]

        self.assertEqual(find_trialjumper_remark_rows(data, messages), {0})

    def test_trialjumper_uses_smallest_row_distance_to_large_gap(self):
        data = [
            fol_row(1 if minute == 0 else None, "W", minute + 1, 3204, 10, minute,
                    gap=120 if minute in (5, 6) else 20)
            for minute in range(10)
        ]
        messages = [msg_row(10, 4, "Trialjumper", status="trialjumper")]

        self.assertEqual(find_trialjumper_remark_rows(data, messages), {5})

    def test_trialjumper_without_large_gap_has_no_remark_target(self):
        data = [
            fol_row(1 if minute == 0 else None, "W", minute + 1, 3204, 10, minute,
                    gap=100)
            for minute in range(6)
        ]
        messages = [msg_row(10, 2, "Trialjumper", status="trialjumper")]

        self.assertEqual(find_trialjumper_remark_rows(data, messages), set())

    def test_bib_message_writes_message_and_gate_in_same_jump_id(self):
        changed = self.sync(
            [msg_row(10, 29, "Gate changed", bib=3, gate=7)]
        )

        self.assertEqual(self.data[3][26], "Gate changed")
        self.assertEqual(self.data[3][25], 7)
        self.assertIn((3, 26), changed)
        self.assertIn((3, 25), changed)

    def test_bib_message_can_write_gate_when_message_is_empty(self):
        changed = self.sync(
            [msg_row(10, 29, None, bib=3, gate=8)]
        )

        self.assertIsNone(self.data[3][26])
        self.assertEqual(self.data[3][25], 8)
        self.assertEqual(changed, {(3, 25)})

    def test_bib_is_not_searched_in_another_jump_id(self):
        changed = self.sync(
            [msg_row(10, 12, "Wrong block", bib=3, gate=8)]
        )

        self.assertIsNone(self.data[3][26])
        self.assertIsNone(self.data[3][25])
        self.assertEqual(changed, set())

    def test_empty_jump_id_is_skipped(self):
        changed = self.sync(
            [msg_row(10, 5, "Unknown Codex", codex=1234)]
        )

        self.assertEqual(changed, set())
        self.assertTrue(all(row[26] is None for row in self.data))

    def test_public_calculation_preserves_inserted_block_structure(self):
        updated, changed = run_calculation(
            self.data,
            [msg_row(10, 16, "Ben Beta not permitted", bib=2, status="NPS")],
        )

        self.assertEqual(updated[2][0:3], [2, "Q", 2])
        self.assertEqual(updated[3][0:2], [None, None])
        self.assertEqual(updated[2][3:6], ["Beta", "Ben", "AUT"])
        self.assertIn((2, 24), changed)


if __name__ == "__main__":
    unittest.main()
