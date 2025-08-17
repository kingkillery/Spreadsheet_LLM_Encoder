import unittest
from unittest.mock import patch
import chain_of_spreadsheet as cos

class TestChainOfSpreadsheet(unittest.TestCase):

    def setUp(self):
        self.sample_encoding = {
            "sheets": {
                "Sheet1": {
                    "cells": {
                        "Header1": ["A1"],
                        "Value1": ["A2"]
                    }
                }
            }
        }
        self.sample_query = "What is the value?"

    @patch('chain_of_spreadsheet._call_llm')
    def test_identify_table(self, mock_call_llm):
        mock_call_llm.return_value = "['range': 'A1:B2']"

        table_range = cos.identify_table(self.sample_encoding, self.sample_query)

        # Check that the LLM was called
        mock_call_llm.assert_called_once()
        # Check that the prompt was formatted correctly
        prompt_arg = mock_call_llm.call_args[0][0]
        self.assertIn(self.sample_query, prompt_arg)
        self.assertIn('"cells": {"Header1": ["A1"]', prompt_arg)
        # Check that the response was parsed correctly
        self.assertEqual(table_range, "A1:B2")

    @patch('chain_of_spreadsheet._call_llm')
    def test_generate_response(self, mock_call_llm):
        mock_call_llm.return_value = "[C5]"
        sheet_data = self.sample_encoding["sheets"]["Sheet1"]

        response = cos.generate_response(sheet_data, self.sample_query)

        mock_call_llm.assert_called_once()
        prompt_arg = mock_call_llm.call_args[0][0]
        self.assertIn(self.sample_query, prompt_arg)
        self.assertEqual(response, "[C5]")

    @patch('chain_of_spreadsheet.generate_response')
    def test_table_split_qa_large_table(self, mock_generate_response):
        # Make the encoding large enough to trigger splitting
        large_encoding = {"cells": {"a" * 5000: ["A1"]}}
        mock_generate_response.return_value = "Sub-answer"

        response = cos.table_split_qa(large_encoding, "A1:Z100", self.sample_query)

        # Check that it tried to generate responses for chunks
        self.assertGreater(mock_generate_response.call_count, 1)
        self.assertIn("Aggregated answers", response)

    @patch('chain_of_spreadsheet.generate_response')
    def test_table_split_qa_small_table(self, mock_generate_response):
        mock_generate_response.return_value = "Direct answer"

        response = cos.table_split_qa(self.sample_encoding["sheets"]["Sheet1"], "A1:A2", self.sample_query)

        # Check that it called generate_response directly without splitting
        mock_generate_response.assert_called_once()
        self.assertEqual(response, "Direct answer")


if __name__ == '__main__':
    unittest.main()
