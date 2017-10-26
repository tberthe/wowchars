import sys, os
import string, random

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "../../")))

from wowchars import SheetConnector, RGBColor

class GoogleSheetsTestHelper(object):
    def __init__(self):
        self.connector = None

    def connect(self, spreadsheet_id, dry_run=False):
        if self.connector:
            print("*WARN* Already connected")
            return
        self.connector = SheetConnector(spreadsheet_id, dry_run)

    def sheet_exists(self, sheet_name):
        return self.connector.sheet_exists(sheet_name)

    def sheet_does_not_exist(self, sheet_name):
        return self.connector.sheet_exists(sheet_name)

    def get_values(self, range_value):
        res =  self.connector.get_values(range_value)
        #print("*WARN* get: ", res)
        return res

    def update_values(self, range_value, values):
        update_data = [{
            "values": values,
            "range": range_value,
        }]
        return self.connector.update_values(update_data)

    def check_or_create_sheet(self, sheet_name):
        return self.connector.check_or_create_sheet(sheet_name)

    def delete_sheet(self, sheet_name):
        self.connector.get_sheets()
        return self.connector.delete_sheet(sheet_name)

    def ensure_headers(self, sheetName, fieldnames):
        self.connector.ensure_headers(sheetName, fieldnames)

    def remove_extra_sheets(self):
        sheets = self.connector.get_sheets()
        for s in sheets:
            sheet_id = sheets[s]
            if sheet_id:
                self.connector.delete_sheet(s)

    def get_random_cell_index(self): 
        c = (random.choice(string.ascii_uppercase), random.randrange(2, 21))
        return c

    def set_cell_color(self, sheetName, cell, hex_color):
        color = RGBColor.from_hex(hex_color)
        self.connector.set_background_color(sheetName, cell[0], cell[1], color)

    def get_cell_hex_color(self, sheetName, cell):
        color = self.connector.get_background_color(sheetName, cell[0], cell[1])
        return color.to_hex()
