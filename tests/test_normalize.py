import unittest
from datetime import datetime

from core.models import IssueCode, Severity
from core.normalize import (
    clean_text,
    has_letters,
    looks_like_legacy_font,
    normalize_account,
    normalize_amount,
    normalize_card,
    normalize_header,
    normalize_name_key,
    remove_diacritics,
    strip_edges,
)


class TestText(unittest.TestCase):
    def test_clean_text(self) -> None:
        self.assertEqual(clean_text("  Nguyễn\u00a0 Văn\n A \u200b"), "Nguyễn Văn A")
        self.assertEqual(clean_text(None), "")
        self.assertEqual(clean_text(12.0), "12")

    def test_header(self) -> None:
        self.assertEqual(normalize_header("Hồng Lâm\nChuyển  Khoản"), normalize_header("hồng lâm chuyển khoản"))
        self.assertNotEqual(normalize_header("Số thẻ"), normalize_header("So the"))  # giữ dấu

    def test_name_key(self) -> None:
        self.assertEqual(remove_diacritics("Đỗ Hữu Thành"), "Do Huu Thanh")
        self.assertEqual(normalize_name_key("  đỗ  hữu thành "), "DO HUU THANH")

    def test_legacy_font(self) -> None:
        self.assertTrue(looks_like_legacy_font("Nguyeãn Thaønh Coâng"))
        self.assertTrue(looks_like_legacy_font("Ñoaøn Thò Tuyeát"))
        self.assertFalse(looks_like_legacy_font("Nguyễn Thành Công"))
        self.assertFalse(looks_like_legacy_font("NGUYEN THANH CONG"))

    def test_strip_edges_keeps_middle(self) -> None:
        self.assertEqual(strip_edges("\u202f\u200b 0587 952 \u00a0"), "0587 952")
        self.assertEqual(strip_edges(None), "")

    def test_has_letters(self) -> None:
        self.assertTrue(has_letters("Đỗ"))
        self.assertFalse(has_letters("12345 678"))
        self.assertFalse(has_letters(""))


class TestCard(unittest.TestCase):
    def test_string_keeps_leading_zero(self) -> None:
        r = normalize_card(" 02101003 ")
        self.assertTrue(r.ok)
        self.assertEqual(r.value, "02101003")

    def test_text_that_looks_like_date_is_kept(self) -> None:
        self.assertEqual(normalize_card("4/7/2012").value, "4/7/2012")

    def test_number_warns(self) -> None:
        r = normalize_card(2101003.0)
        self.assertEqual(r.value, "2101003")
        self.assertEqual(r.code, IssueCode.CARD_STORED_AS_NUMBER)
        self.assertEqual(r.severity, Severity.WARNING)

    def test_negative_or_fraction_is_error(self) -> None:
        self.assertTrue(normalize_card(-5).is_error)
        self.assertTrue(normalize_card(12.5).is_error)
        self.assertTrue(normalize_card(True).is_error)

    def test_date_is_error(self) -> None:
        r = normalize_card(datetime(2012, 4, 7))
        self.assertTrue(r.is_error)
        self.assertIsNone(r.value)

    def test_empty(self) -> None:
        self.assertTrue(normalize_card(None).is_empty)
        self.assertTrue(normalize_card("   ").is_empty)


class TestAccount(unittest.TestCase):
    def test_text_with_leading_zero(self) -> None:
        self.assertEqual(normalize_account("0587952606").value, "0587952606")

    def test_int_and_float(self) -> None:
        self.assertEqual(normalize_account(24349081).value, "24349081")
        self.assertEqual(normalize_account(107886464707.0).value, "107886464707")

    def test_float_with_fraction_is_error(self) -> None:
        self.assertTrue(normalize_account(24349081.5).is_error)

    def test_float_too_long_is_error(self) -> None:
        self.assertTrue(normalize_account(1234567890123456.0).is_error)

    def test_internal_space_is_error_not_fixed(self) -> None:
        r = normalize_account("0587 952 606")
        self.assertTrue(r.is_error)
        self.assertIsNone(r.value)

    def test_letters_and_dots_are_error(self) -> None:
        self.assertTrue(normalize_account("1.07886E+11").is_error)
        self.assertTrue(normalize_account("ACB123456").is_error)

    def test_unusual_length_warns(self) -> None:
        r = normalize_account("12345")
        self.assertEqual(r.value, "12345")
        self.assertEqual(r.code, IssueCode.ACCOUNT_UNUSUAL_LENGTH)

    def test_bool_date_error(self) -> None:
        self.assertTrue(normalize_account(True).is_error)
        self.assertTrue(normalize_account(datetime(2020, 1, 1)).is_error)

    def test_empty(self) -> None:
        self.assertTrue(normalize_account(None).is_empty)
        self.assertTrue(normalize_account("").is_empty)


class TestAmount(unittest.TestCase):
    def test_int_float_text(self) -> None:
        self.assertEqual(normalize_amount(5917028).value, 5917028)
        self.assertEqual(normalize_amount(5917028.0).value, 5917028)
        self.assertEqual(normalize_amount(" 5917028 ").value, 5917028)

    def test_float_noise_accepted(self) -> None:
        self.assertEqual(normalize_amount(5917028.0000000001).value, 5917028)

    def test_fraction_is_error_no_rounding(self) -> None:
        r = normalize_amount(5917028.11)
        self.assertEqual(r.code, IssueCode.AMOUNT_NOT_INTEGER)
        self.assertIsNone(r.value)

    def test_separator_text_is_error(self) -> None:
        self.assertEqual(normalize_amount("5.917.028").code, IssueCode.AMOUNT_NOT_NUMBER)
        self.assertEqual(normalize_amount("5,917,028").code, IssueCode.AMOUNT_NOT_NUMBER)

    def test_negative_zero_large(self) -> None:
        self.assertEqual(normalize_amount(-1).code, IssueCode.AMOUNT_NEGATIVE)
        self.assertTrue(normalize_amount(-1).is_error)
        self.assertEqual(normalize_amount(0).code, IssueCode.AMOUNT_ZERO)
        self.assertEqual(normalize_amount(0).severity, Severity.WARNING)
        self.assertEqual(normalize_amount(2_000_000_000).code, IssueCode.AMOUNT_UNUSUALLY_LARGE)

    def test_bool_date_error(self) -> None:
        self.assertTrue(normalize_amount(True).is_error)
        self.assertTrue(normalize_amount(datetime(2020, 1, 1)).is_error)

    def test_empty(self) -> None:
        self.assertTrue(normalize_amount(None).is_empty)


if __name__ == "__main__":
    unittest.main()
