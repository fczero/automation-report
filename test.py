#!/usr/bin/env python3

import unittest
import ccp_daily_automation as script

from openpyxl import Workbook
from openpyxl.styles import colors
from openpyxl.styles import Color, PatternFill, Font, Border, NamedStyle
from openpyxl.styles import Alignment, Side
from openpyxl.formatting import Rule
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter, rows_from_range
from openpyxl.utils import units
from openpyxl.worksheet.datavalidation import DataValidation
from selenium import webdriver as wd
import json
import requests


class testScript(unittest.TestCase):

    def setUp(self):
        pass

    def test_arghandler(self):
#        unittest.mock.patch('sys.argv',)
#        script.arghandler()
        pass

    def test_prefix_scenario(self):
        input = 'Scenario:This is a test scenario'
        expected = 'This is a test scenario'
        actual = script.fixPrefix(input)
        self.assertEqual(expected, actual)

    def test_prefix_background(self):
        input = 'Background:'
        expected = 'Background Step'
        actual = script.fixPrefix(input)
        self.assertEqual(expected, actual)

    def test_remove_underscore(self):
        input = 'left_right'
        expected = 'leftright'
        actual = script.remUnderscore(input)
        self.assertEqual(expected, actual)

    def test_scrape_duration(self):
        input = '9 mins and 44 secs and 812 ms'
        expected = '9mi 44s 812ms'
        actual = script.duration(input)
        self.assertEqual(expected, actual)
        input = '1 hour and 812 ms'
        expected = '1h 812ms'
        actual = script.duration(input)
        self.assertEqual(expected, actual)

if __name__ == "__main__":
    unittest.main()
