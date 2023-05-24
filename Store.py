from tkinter import *
from tkinter import filedialog
import customtkinter
import os
import mysql.connector
import pandas as pd
from pandas.io import sql as sql
from pyxlsb import open_workbook as open_xlsb
import openpyxl
from pyepc import SGTIN
from pyepc.exceptions import DecodingError


class Store:

    def __init__(self, store_num, date_input, cycle, cycle_output, item_file, qb_path, matching, qb_matching,
                 total_items, repl_nbr, expected, combined, combined_qb_matching, combined_repl, UE, DU, UU, errorEPCs,
                 errorMessages):
        self.store_num = store_num
        self.date_input = date_input
        self.cycle = cycle
        self.cycle_output = cycle_output
        self.item_file = item_file
        self.qb_path = qb_path
        self.matching = matching
        self.qb_matching = qb_matching
        self.total_items = total_items
        self.repl_nbr = repl_nbr
        self.expected = expected
        self.combined = combined
        self.combined_qb_matching = combined_qb_matching
        self.combined_repl = combined_repl
        self.UE = UE
        self.DU = DU
        self.UU = UU
        self.errorEPCs = errorEPCs
        self.errorMessages = errorMessages

    def set_cycle(self, cycle_path):
        self.cycle = cycle_path

    def set_cycle_output(self, cycle_output_path):
        self.cycle_output = cycle_output_path

    def set_item_file(self, item_file_path):
        self.item_file = item_file_path

    def set_qb_path(self, qb_path):
        self.qb_path = qb_path

    def set_matching(self, matching_df):
        self.matching = matching_df

    def set_qb_matching(self, qb_matching_df):
        self.qb_matching = qb_matching_df

    def set_total_items(self, total_items_df):
        self.total_items = total_items_df

    def set_repl_nbr(self, repl_nbr):
        self.repl_nbr = repl_nbr

    def set_expected(self, expected_df):
        self.expected = expected_df

    def set_combined(self, combined):
        self.combined = combined

    def set_qb_combined(self, qb_combined):
        self.combined_qb_matching = qb_combined

    def set_combined_repl(self, combined_repl):
        self.combined_repl = combined_repl

    def set_store_num(self, store_number):
        self.store_num = store_number

    def set_date_input(self, inputted_date):
        self.date_input = inputted_date

    def set_UE(self, UE):
        self.UE = UE

    def set_UU(self, UU):
        self.UU = UU

    def set_DU(self, DU):
        self.DU = DU

    def set_error_EPCs(self, errorEPCs):
        self.errorEPCs = errorEPCs

    def set_error_messages(self, errorMessages):
        self.errorMessages = errorMessages

    def get_cycle(self):
        return self.cycle

    def get_cycle_output(self):
        return self.cycle_output

    def get_item_file(self):
        return self.item_file

    def get_qb_path(self):
        return self.qb_path

    def get_matching(self):
        return self.matching

    def get_qb_matching(self):
        return self.qb_matching

    def get_total_items(self):
        return self.total_items

    def get_repl_nbr(self):
        return self.repl_nbr

    def get_expected(self):
        return self.expected

    def get_combined(self):
        return self.combined

    def get_qb_combined(self):
        return self.combined_qb_matching

    def get_combined_repl(self):
        return self.combined_repl

    def get_store_num(self):
        return self.store_num

    def get_date_input(self):
        return self.date_input

    def get_UE(self):
        return self.UE

    def get_UU(self):
        return self.UU

    def get_DU(self):
        return self.DU

    def get_error_EPCs(self):
        return self.errorEPCs

    def get_error_messages(self):
        return self.errorMessages



    def toString(self):
        string = "Store Number: " + str(self.get_store_num()) \
                 + "\n\tDate: " + str(self.get_date_input()) \
                 + "\n\tCycle Count path: " + str(self.get_cycle()) \
                 + "\n\tCycle Count Output path: " + str(self.get_cycle_output()) \
                 + "\n\tItem File paths: " + str(self.get_item_file()) \
        + "\n\tQB Master Items path: " + str(self.get_qb_path()) \
        + "\n\tMatching Data Frame: " + str(self.get_matching()) \
        + "\n\tQB Matching Data Frame: " + str(self.get_qb_matching()) \
        + "\n\tTotal Items Data Frame: " + str(self.get_total_items()) \
        + "\n\tREPL_GROUP_NBR Data Frame: " + str(self.get_repl_nbr()) \
        + "\n\tExpected Items Data Frame: " + str(self.get_expected()) \
        + "\n\tCombined Items Data Frame: " + str(self.get_combined()) \
        + "\n\tCombined QB Matching Data Frame: " + str(self.get_qb_combined()) \
        + "\n\tCombined REPL Data Frame: " + str(self.get_combined_repl())
        return string
