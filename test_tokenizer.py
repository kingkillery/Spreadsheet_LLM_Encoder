"""Tests for tokenizer.py — tiktoken wrapper and char-based fallback."""
import os
import sys
import unittest
from unittest.mock import patch

sys.path.insert(0, os.path.dirname(__file__))
import tokenizer as tok


class TestCountTokensEmptyString(unittest.TestCase):

    def test_empty_string_returns_zero(self):
        result = tok.count_tokens("")
        self.assertEqual(result, 0)


class TestCountTokensPositiveResult(unittest.TestCase):

    def test_non_empty_string_returns_positive_int(self):
        result = tok.count_tokens("hello world")
        self.assertIsInstance(result, int)
        self.assertGreater(result, 0)


class TestTiktokenAvailability(unittest.TestCase):

    def test_is_tiktoken_available_returns_bool(self):
        result = tok.is_tiktoken_available()
        self.assertIsInstance(result, bool)


class TestTiktokenWhenAvailable(unittest.TestCase):

    def test_two_different_texts_give_different_counts_when_tiktoken_available(self):
        if not tok.is_tiktoken_available():
            self.skipTest("tiktoken not installed")
        short = tok.count_tokens("hi")
        long_ = tok.count_tokens("The quick brown fox jumps over the lazy dog " * 10)
        self.assertNotEqual(short, long_)


class TestFallbackBehavior(unittest.TestCase):

    def test_fallback_returns_floor_div_4_for_non_empty(self):
        text = "abcdefghijklmnop"  # 16 chars -> 16//4 = 4
        with patch.object(tok, "_TIKTOKEN_AVAILABLE", False):
            tok._FALLBACK_WARNED = False
            result = tok.count_tokens(text)
        self.assertEqual(result, len(text) // 4)

    def test_fallback_returns_zero_for_empty_string(self):
        with patch.object(tok, "_TIKTOKEN_AVAILABLE", False):
            tok._FALLBACK_WARNED = False
            result = tok.count_tokens("")
        self.assertEqual(result, 0)

    def test_fallback_returns_at_least_one_for_short_text(self):
        # "ab" -> 2//4 = 0 but max(1, 0) = 1
        with patch.object(tok, "_TIKTOKEN_AVAILABLE", False):
            tok._FALLBACK_WARNED = False
            result = tok.count_tokens("ab")
        self.assertEqual(result, 1)

    def test_fallback_count_scales_with_length(self):
        text_8 = "abcdefgh"   # 8 chars -> 2
        text_16 = "abcdefgh" * 2  # 16 chars -> 4
        with patch.object(tok, "_TIKTOKEN_AVAILABLE", False):
            tok._FALLBACK_WARNED = False
            count_8 = tok.count_tokens(text_8)
            count_16 = tok.count_tokens(text_16)
        self.assertLess(count_8, count_16)


if __name__ == "__main__":
    unittest.main()
