from column_utils import column_index, column_letter
from rgb_color import RGBColor

import httplib2
import logging
import os
import re

# import googleapiclient
from apiclient import discovery
from oauth2client import client, tools
from oauth2client.file import Storage

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'wowchars'


class GoogleSheetsConnector:
    """Helper class to use Google Sheets"""
    def __init__(self, sheet_id, dry_run):
        """Constructor

        Args:
            sheet_id (str): ID of the document.
            dry_run (bool): if True, do not modify the document
        """
        self.dry_run = dry_run
        self.credentials = self.get_credentials()

        http = self.credentials.authorize(httplib2.Http())
        discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                        'version=v4')
        self.service = discovery.build('sheets', 'v4', http=http,
                                       discoveryServiceUrl=discoveryUrl,
                                       cache_discovery=False)

        self.spreadsheetId = sheet_id

    def check_or_create_sheet(self, sheet_name):
        """Check if the sheet exists in the document, create it otherwise

        Args:
            sheet_name (str): name of the sheet to check
        """
        if self.sheet_exists(sheet_name):
            return

        body = {
            "requests": [
                {
                    "addSheet": {
                        "properties": {
                            "title": sheet_name,
                            # "gridProperties": {
                            #   "rowCount": 20,
                            #   "columnCount": 12
                            # },
                            # "tabColor": {
                            #   "red": 1.0,
                            #   "green": 0.3,
                            #   "blue": 0.4
                            # }
                        }
                    }
                }
            ]
        }
        self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheetId, body=body).execute()

    def sheet_exists(self, sheet_name):
        """Check if the sheet exists in the document

        Args:
            sheet_name (str): name of the sheet to check
        """
        return sheet_name in self.get_sheets()

    def delete_sheet(self, sheet_name):
        """Delete the sheet in the spreadsheet

        Args:
            sheet_name (str): name of the sheet to delete
        """
        sheets = self.get_sheets()
        if sheet_name not in sheets:
            return

        body = {
            "requests": [
                {
                    "deleteSheet": {
                        "sheetId": sheets[sheet_name]
                    }
                }
            ]
        }
        self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheetId, body=body).execute()

    def get_sheets(self):
        """Get the sheets in the doc

        Returns:
            (dict) keys are names, values are IDs
        """
        sheet_metadata = self.service.spreadsheets().get(spreadsheetId=self.spreadsheetId).execute()
        sheets = sheet_metadata.get('sheets', '')
        return {s["properties"]["title"]: s["properties"]["sheetId"] for s in sheets}

    def ensure_headers(self, sheet_name, fieldnames):
        """Ensure the sheet exists and contains the specified headers

        Args:
            sheet_name (str): name of the sheet to check
            sheet_name (str array): headers to check
        """
        self.check_or_create_sheet(sheet_name)
        values, _range = self.get_values(sheet_name + "!1:1")
        g_headers = values[0] if values else []

        appended_headers = []

        # building indexes map and adding extra headers
        g_headers_indexes = {g_h: i for i, g_h in enumerate(g_headers)}
        for field in fieldnames:
            try:
                g_headers_indexes[field] = g_headers.index(field)
            except ValueError:
                g_headers.append(field)
                appended_headers.append(field)
                g_headers_indexes[field] = len(g_headers)

        if(appended_headers):
            logging.info("Adding headers %s", appended_headers)
            update_data = SheetBatchUpdateData()
            update_data.add_data(sheet_name,
                                 column_letter(len(g_headers_indexes) - len(appended_headers)), 1,
                                 column_letter(len(g_headers_indexes) - 1), 1,
                                 [appended_headers])
            self.update_values(update_data)

    def update_values(self, update_data):
        """Update values in the document

        Args:
            update_data (array): data to update, ex: [{"values": [["val1", "val2"]], "range": "sheet2!A2:B2"}]
        """
        self.ensure_no_missing_row_or_column(update_data)

        data = update_data.to_query_data()
        body = {"data": data, "value_input_option": "USER_ENTERED"}
        logging.info("%sUpdating data in Google sheets: %s", ("DRYRUN: " if self.dry_run else ""), data)
        if not self.dry_run:
            self.service.spreadsheets().values().batchUpdate(spreadsheetId=self.spreadsheetId, body=body).execute()

    def ensure_no_missing_row_or_column(self, update_data):
        """Add rows or colums if needed.

        Args:
            update_data (array): data to update, ex: [{"values": [["val1", "val2"]], "range": "sheet2!A2:B2"}]
        """
        max_cols = update_data.get_max_columns()

        for mc in max_cols:
            last_col = self.get_last_column(mc)
            last_col_index = column_index(last_col)
            max_col_index = column_index(max_cols[mc])
            if max_col_index > last_col_index:
                sheet_id = self.get_sheet_id(mc)
                self.append_columns(sheet_id, max_col_index - last_col_index)

        max_rows = update_data.get_max_rows()
        for mr in max_rows:
            last_row_index = self.get_last_row(mr)
            max_row_index = max_rows[mr]
            if max_row_index > last_row_index:
                sheet_id = self.get_sheet_id(mr)
                self.append_rows(sheet_id, max_row_index - last_row_index)

    def get_values(self, rangeName):
        """Get values

        Args:
            rangeName (str): range of the values to get. Ex: "sheet2!A2:B2"
        """
        result = self.service.spreadsheets().values().get(
            spreadsheetId=self.spreadsheetId, range=rangeName).execute()
        return result.get('values', []), result.get('range', None)

    def get_credentials(self, flags=None):
        """Gets valid user credentials from storage.

        If nothing has been stored, or if the stored credentials are invalid,
        the OAuth2 flow is completed to obtain the new credentials.

        Returns:
            Credentials, the obtained credential.
        """
        home_dir = os.path.expanduser('~')
        credential_dir = os.path.join(home_dir, '.credentials')
        if not os.path.exists(credential_dir):
            os.makedirs(credential_dir)
        credential_path = os.path.join(credential_dir,
                                       'sheets.googleapis.com-python-wowchars.json')

        store = Storage(credential_path)
        credentials = store.get()
        if not credentials or credentials.invalid:
            flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
            flow.user_agent = APPLICATION_NAME
            if flags:
                credentials = tools.run_flow(flow, store, flags)
            else:  # Needed only for compatibility with Python 2.6
                credentials = tools.run(flow, store)
            print('Storing credentials to ' + credential_path)
        return credentials

    def get_background_color(self, sheet, column, row):
        """Get background color of a cell

        Args:
            sheet (string): name of the sheet
            column (string): column id of the cell
            row (int): row of the cell

        Returns:
            The background color as a RGBColor object
        """
        ranges = "%s!%s%d" % (sheet, column, row)
        data = self.service.spreadsheets().get(spreadsheetId=self.spreadsheetId, ranges=ranges, includeGridData=True).execute()
        v = data["sheets"][0]["data"][0]["rowData"][0]["values"][0]
        if "effectiveFormat" in v:
            return RGBColor.from_float_rgb_dict(v["effectiveFormat"]["backgroundColor"])
        else:
            return RGBColor.from_float_rgb_dict(v["userEnteredFormat"]["backgroundColor"])

    def set_background_color(self, sheet_name, column, row, rgb_color):
        """Set the background color of a cell

        Args:
            sheet (string): name of the sheet
            column (string): column id of the cell
            row (int): row of the cell
        """
        col_i = column_index(column) if type(column) is str else column

        sheets = self.get_sheets()

        body = {
            "requests": [
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": sheets[sheet_name],
                            "startRowIndex": row - 1,
                            "endRowIndex": row,
                            "startColumnIndex": col_i,
                            "endColumnIndex": col_i + 1
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": rgb_color.to_float_rgb_dict()
                            }
                        },
                        "fields": "userEnteredFormat(backgroundColor)"
                    }
                },
            ]
        }
        self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheetId, body=body).execute()

    def get_sheet_id(self, sheet_name):
        """Get the ID of a sheet

        Args:
            sheet (string): name of the sheet

        Returns:
            The ID of the sheet
        """
        spreadsheet = self.service.spreadsheets().get(spreadsheetId=self.spreadsheetId).execute()
        for _sheet in spreadsheet['sheets']:
            if _sheet['properties']['title'] == sheet_name:
                return _sheet['properties']['sheetId']
        return None

    def append_columns(self, sheet_id, nb_cols):
        """Add columns at the right of a sheet

        Args:
            sheet_id (string): ID of the sheet
            nb_cols(int): number of columns to add
        """
        logging.debug("Adding %d column(s)" % nb_cols)
        request_body = {
            "requests": [{
                "appendDimension": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "length": nb_cols
                }
            }]
        }
        self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheetId, body=request_body).execute()

    def append_rows(self, sheet_id, nb_rows):
        """Add rows at the end of a sheet

        Args:
            sheet_id (string): ID of the sheet
            nb_rows(int): number of rows to add
        """
        logging.debug("Adding %d row(s)" % nb_rows)
        request_body = {
            "requests": [{
                "appendDimension": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "length": nb_rows
                }
            }]
        }
        self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheetId, body=request_body).execute()

    def get_last_column(self, sheet_name):
        """Get ID of the last column

        Args:
            sheet_name (string): name of the sheet

        Returns:
            The ID of the last column (ex "AF")
        """
        first_line_range = "{sheet}!1:1".format(sheet=sheet_name)
        _, res_range = self.get_values(first_line_range)
        match = re.match(r"(?P<sheet>.*)!(?P<start_col>\w+)(\d+)(:(?P<end_col>\w+)(\d+))?", res_range)
        return match.group("end_col") or match.group("start_col")

    def get_last_row(self, sheet_name):
        """Get number of the last row

        Args:
            sheet_name (string): name of the sheet

        Returns:
            The number of the last row
        """
        first_row_range = "{sheet}!A:A".format(sheet=sheet_name)
        _, res_range = self.get_values(first_row_range)
        match = re.match(r"(?P<sheet>.*)!(\w+)(?P<start_row>\d+)(:(\w+)(?P<end_row>\d+))?", res_range)
        return int(match.group("end_row") or match.group("start_row"))


class SheetSingleUpdateData:
    """Classe defining a single set of data update for a sheet"""

    def __init__(self, sheet_name, start_col, start_row, end_col, end_row, values):
        """Constructor.

        Args:
            sheet_name (string): name of the sheet
            start_col (string): starting column of the data
            start_row (int): starting row of the data
            end_col (string): ending column of the data
            start_row (int): starting row of the data
            values: values to set as an array (1D) or array of array (2D)
        """
        self.sheet_name = sheet_name
        self.start_col = start_col
        self.end_col = end_col
        self.start_row = start_row
        self.end_row = end_row
        self.values = values

    def to_query_data(self):
        """Get the update data in the google sheet query format

        Returns:
            The update as an object useable by the google sheet API
        """
        return {
            "values": self.values,
            "range": "%s!%s%d:%s%d" % (self.sheet_name, self.start_col, self.start_row, self.end_col, self.end_row)
        }


class SheetBatchUpdateData(list):
    """Classe defining a batch of sets of data update for a sheet"""

    def add_data(self, sheet_name, start_col, start_row, end_col, end_row, values):
        """Add a single set of data to update.

        Args:
            sheet_name (string): name of the sheet
            start_col (string): starting column of the data
            start_row (int): starting row of the data
            end_col (string): ending column of the data
            start_row (int): starting row of the data
            values: values to set as an array (1D) or array of array (2D)
        """
        self.append(SheetSingleUpdateData(sheet_name, start_col, start_row, end_col, end_row, values))

    def get_max_columns(self):
        """Get max updated column for each sheet

        Returns:
            a dictionnary  {sheet name: max column}
        """
        # grabbing max ranges for each sheet
        max_columns = {}     # {sheet: max column}
        for data in self:
            if data.sheet_name not in max_columns:
                max_columns[data.sheet_name] = data.end_col
            elif data.end_col > max_columns[data.sheet_name]:
                max_columns[data.sheet_name] = data.end_col
        return max_columns

    def get_max_rows(self):
        """Get max updated row for each sheet

        Returns:
            a dictionnary  {sheet name: max row}
        """
        max_rows = {}  # {sheet: max row}
        for data in self:
            if data.sheet_name not in max_rows:
                max_rows[data.sheet_name] = data.end_row
            elif data.end_row > max_rows[data.sheet_name]:
                max_rows[data.sheet_name] = data.end_row
        return max_rows

    def to_query_data(self):
        """Get the update data in the google sheet query format
        
        Returns:
            The update as an object useable by the google sheet API
        """
        return [d.to_query_data() for d in self]
