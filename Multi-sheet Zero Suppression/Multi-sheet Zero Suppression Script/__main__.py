############################################################################

# Copyright 2022 Workiva Inc.

# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at

#     http://www.apache.org/licenses/LICENSE-2.0

# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

############################################################################


import json
import decimal
import os
import logging
from dataclasses import dataclass
from datetime import datetime

import requests

AUTH_URL = "https://api.sandbox.wdesk.com/iam/v1/oauth2/token"
SS_API_URL = 'https://api.sandbox.wdesk.com/platform/v1/spreadsheets/'

NumberPrecision = {
    'BASIS POINTS': 0.0001,
    'HUNDREDTHS': 0.01,
    'ONES': 1,
    'THOUSANDS': 1000,
    'TEN THOUSANDS': 10_000,
    'MILLIONS': 1_000_000,
    'HUNDRED MILLIONS': 100_000_000,
    'BILLIONS': 1_000_000_000,
    'TRILLIONS': 1_000_000_000_000
}

class ApiAuth:
    def __init__(self, client_id, client_secret, auth_url):
        self._headers = {'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8'}
        self.client_id = client_id
        self.client_secret = client_secret
        self.auth_url = auth_url

    def get_auth_token(self):
        data = 'client_id=' + self.client_id + '&client_secret=' + self.client_secret + '&grant_type=client_credentials'
        token_res = requests.post(self.auth_url, data=data, headers=self._headers)
        print("Auth response: ", token_res)
        token_response = json.loads(token_res.text)
        return token_response['access_token']

@dataclass
class MultiDocumentRow:
    workspace_id: str
    document_id: str
    hide_zeros: str
    client_id: str
    client_secret: str

class MultiDocument:
    def __init__(self, client_id, client_secret, workspace_id, document_id, sheet_id):
        self.api_auth = ApiAuth(client_id, client_secret, AUTH_URL)
        self._accessToken = 'Bearer ' + self.api_auth.get_auth_token()
        self.spreadsheet_api = SpreadsheetApi(self.api_auth.get_auth_token())
        self.workspace_id = workspace_id
        self.document_id = document_id
        self.sheet_id = sheet_id
        self.rows = []

    def run(self):
        table_data = self.spreadsheet_api.get_table_data(self.document_id, self.sheet_id, 'A:E')
        for row in table_data[1:]:
            row_values = []
            empty = True
            for cell in row:
                cell_value = cell['calculatedValue']
                if cell_value != '':
                    empty = False
                row_values.append(cell_value)
            if empty:
                break
            self.rows.append(MultiDocumentRow(*row_values))

        first_row = self.rows[0]
        if first_row.workspace_id == '':
            first_row.workspace_id = self.workspace_id
        if first_row.client_id == '':
            first_row.client_id = self.api_auth.client_id
        if first_row.client_secret == '':
            first_row.client_secret = self.api_auth.client_secret
        for i, row in enumerate(self.rows[1:], start=1):
            if row.workspace_id == '':
                row.workspace_id = self.rows[i - 1].workspace_id
            if row.client_id == '':
                row.client_id = self.rows[i - 1].client_id
            if row.client_secret == '':
                row.client_secret = self.rows[i - 1].client_secret

        for i, row in enumerate(self.rows, start=1):
            auth_token = ApiAuth(row.client_id, row.client_secret, AUTH_URL).get_auth_token()
            spreadsheet_api = SpreadsheetApi(auth_token)

            if row.hide_zeros.casefold() == 'yes':
                output = "Done!"
                try:
                    spreadsheet_api.hide_rows(row.document_id)
                except requests.HTTPError:
                    output = "Error!"

                edit_cells = {
                    "editCells": {
                        "cells": [
                            {
                                "row": i,
                                "column": 5,
                                "value": output
                            },
                            {
                                "row": i,
                                "column": 6,
                                "value": str(datetime.now())
                            }
                        ]
                    }
                }
                url = SS_API_URL + self.document_id + "/sheets/" + self.sheet_id + "/update"
                response = requests.post(url, headers={'Authorization': self._accessToken}, json=edit_cells)
                if response.status_code // 100 != 2:
                    raise requests.HTTPError(f"Error updating top level spreadsheet: {response.text}")

class SpreadsheetApi:
    def __init__(self, access_token):
        self._accessToken = 'Bearer ' + access_token
        self._totalRowsHidden = 0

    def get_document_tables(self, doc_id):
        """
        Retrieves a list of identifiers for all the tables contained in a document.
        """
        ids = []
        url = SS_API_URL + doc_id + '/sheets'
        response = requests.get(url, headers={'Authorization': self._accessToken})
        if response.status_code // 100 != 2:
            raise requests.HTTPError(f"Error getting document tables: {response.text}")
        tables = json.loads(response.text)['data']
        for table in tables:
            ids.append(table['id'])
        return ids

    def get_table_data(self, doc_id, table_id, cell_range=None):
        """
        Retrieves the data for a table.
        """
        url = SS_API_URL + doc_id + "/sheets/" + table_id + "/sheetdata"
        params = {'$fields': 'cells.calculatedValue,cells.effectiveFormats.valueFormat.shownIn,'
                             'cells.effectiveFormats.valueFormat.precision.auto,'
                             'cells.effectiveFormats.valueFormat.precision.value'}
        if cell_range is not None:
            params['$cellrange'] = cell_range
        response = requests.get(url, headers={'Authorization': self._accessToken}, params=params)
        if response.status_code // 100 != 2:
            raise requests.HTTPError(f"Error getting table data: {response.text}")
        return json.loads(response.text, parse_float=decimal.Decimal)['data']['cells']

    def hide_table_rows(self, doc_id, table_id, row_indices):
        """
        Makes a request to hide the rows with indices specified in row_indices.
        """
        if len(row_indices) == 0:
            return

        self._totalRowsHidden += len(row_indices)

        row_indices.sort()

        intervals = []
        start_index = row_indices[0]
        end_index = row_indices[0]

        for index in row_indices:
            if index > end_index + 1:
                interval = {'start': start_index, 'end': end_index}
                intervals.append(interval)
                start_index = index
            end_index = index

        interval = {'start': start_index, 'end': end_index}
        intervals.append(interval)

        url = SS_API_URL + doc_id + "/sheets/" + table_id + "/update"
        response = requests.post(url, headers={'Authorization': self._accessToken},
                                 json={'hideRows': {'intervals': intervals}})
        if response.status_code // 100 != 2:
            raise requests.HTTPError(f"Error hiding table rows: {response.text}")

    def unhide_table_rows(self, doc_id, table_id):
        """
        Makes a request to unhide all the rows within a table.
        """
        # By not specifying start and end they become infinite
        infinite_interval = {}
        intervals = {'unhideRows': {'intervals': [infinite_interval]}}

        url = SS_API_URL + doc_id + "/sheets/" + table_id + "/update"
        response = requests.post(url, headers={'Authorization': self._accessToken}, json=intervals)
        if response.status_code // 100 != 2:
            raise requests.HTTPError(f"Error unhiding table rows: {response.text}")

    def get_rows_as_displayed(self, doc_id, table_id):
        """
        Uses information from table data to create a list of rows with rounded and scaled display values.
        """
        try:
            table_data = self.get_table_data(doc_id, table_id)
        except requests.HTTPError as e:
            raise e

        rows_as_displayed = []
        for row in table_data:
            row_as_displayed = []
            for cell in row:
                calculated_value = cell['calculatedValue']
                if not isinstance(calculated_value, decimal.Decimal):
                    try:
                        float(calculated_value)
                        calculated_value = decimal.Decimal(calculated_value)
                    except ValueError:
                        pass

                displayed_value = calculated_value

                if type(displayed_value) is decimal.Decimal:
                    shown_in = cell['effectiveFormats']['valueFormat']['shownIn']
                    if shown_in is not None:
                        shown_in_value = NumberPrecision[shown_in]
                        displayed_value /= shown_in_value

                    precision = cell['effectiveFormats']['valueFormat']['precision']
                    if precision is not None and not precision['auto']:
                        displayed_value = displayed_value.quantize(decimal.Decimal(10)**precision['value'],
                                                                   decimal.ROUND_HALF_UP)

                row_as_displayed.append(displayed_value)
            rows_as_displayed.append(row_as_displayed)

        return rows_as_displayed

    def section_rows_to_hide(self, start_row, stop_row, zero_rows, has_numeric_data, has_non_zero_numeric_data):
        """
        Creates a list of row indicies to hide for a content section.
        """
        if has_non_zero_numeric_data:
            return zero_rows
        elif has_numeric_data:
            return range(start_row, stop_row + 1)

        # Don't hide sections with no numeric rows.
        return []

    def find_rows_to_hide(self, rows):
        """
        Creates a list of row indices that should be hidden, determined by whether the row is a zero row
        or part of a content section that consists entirely of zero rows.
        """
        rows_to_hide = []

        # Content section variables
        title_row = None
        zero_rows = []
        has_numeric_data = False
        has_non_zero_numeric_data = False

        i = 0
        for i, row in enumerate(rows):

            # Row variables
            is_spacer_row = True
            has_nums = False
            all_zeroes = True

            for cell in row:
                if cell != '':
                    is_spacer_row = False
                if type(cell) is decimal.Decimal:
                    has_nums = True
                    if cell != 0:  # This row has numbers, but is not a zero row
                        all_zeroes = False
                        break

            if is_spacer_row is True:
                if title_row is not None:
                    rows_to_hide.extend(
                        self.section_rows_to_hide(title_row, i, zero_rows, has_numeric_data, has_non_zero_numeric_data))
                    title_row = None
                    zero_rows = []
                    has_numeric_data = False
                    has_non_zero_numeric_data = False
            else:
                if title_row is None:
                    title_row = i
                if has_nums:
                    has_numeric_data = True
                    if all_zeroes:  # zero row
                        zero_rows.append(i)
                    else:
                        has_non_zero_numeric_data = True

        if title_row is not None:
            rows_to_hide.extend(
                self.section_rows_to_hide(title_row, i, zero_rows, has_numeric_data, has_non_zero_numeric_data))

        return rows_to_hide

    def hide_rows(self, doc_id, table_ids=None):
        """
        Hides all the zero rows and empty content sections in every table in the document specified by
        the document identifier.
        """
        if table_ids is None:
            try:
                table_ids = self.get_document_tables(doc_id)
            except requests.HTTPError as e:
                logging.exception(e)
                raise e
        print(f"Hiding rows in {len(table_ids)} tables in document with id: {doc_id}")
        for i, table_id in enumerate(table_ids):
            print(f"Hiding rows in table {i + 1}: {table_id}")
            try:
                rows_as_displayed = self.get_rows_as_displayed(doc_id, table_id)
            except requests.HTTPError as e:
                logging.exception(e)
                raise e
            rows_to_hide = self.find_rows_to_hide(rows_as_displayed)
            try:
                self.hide_table_rows(doc_id, table_id, rows_to_hide)
            except requests.HTTPError as e:
                logging.exception(e)
                raise e
        print(f"Total rows hidden: {self._totalRowsHidden}")

    def unhide_all_rows(self, doc_id):
        """
        Unhides every row in every table in the document specified by the document identifier.
        """
        try:
            table_ids = self.get_document_tables(doc_id)
        except requests.HTTPError as e:
            logging.exception(e)
            raise e
        print(f"Unhiding {len(table_ids)} tables in document with id: {doc_id}")
        for i, table_id in enumerate(table_ids):
            print(f"Unhiding table {i + 1}: {table_id}")
            try:
                self.unhide_table_rows(doc_id, table_id)
            except requests.HTTPError as e:
                logging.exception(e)
                raise e

def getSpreadsheetId(wurl): 
    
    start = wurl.find("sheets_") 
    
    if (start >= 0):
        return wurl[start + len("sheets_"):] 
        
    return None

def getSpreadsheetSectionId(sheetId, wurl): 
    
    start = wurl.find("sheets_" + sheetId + "_") 

    if (start >= 0):
        return wurl[start + len("sheets_" + sheetId + "_"):] 
        
    return None

def main():
    os.environ.update(
        {
            "WORKIVA_CLIENT_ID": "client_id_here",
            "WORKIVA_CLIENT_SECRET": "client_secret_here",
        }
    )

    #CLIENT_ID = os.getenv('CLIENT_ID')
    #CLIENT_SECRET = os.getenv('CLIENT_SECRET')
    #DOCUMENT_ID = os.getenv('DOCUMENT_ID')

    CLIENT_ID = os.getenv('CLIENT_ID') # Hardcode this if triggered from Integrated Automation
    CLIENT_SECRET = os.getenv('CLIENT_SECRET') # Hardcode this if triggered from Integrated Automation

    WORKSPACE_ID = ""

    DOCUMENT_ID = getSpreadsheetId(os.getenv('INPUT_SHEET_ID'))
    
    if DOCUMENT_ID != None :
        SHEET_ID = getSpreadsheetSectionId(DOCUMENT_ID, os.getenv('INPUT_RESOURCE_ID'))

    auth_token = ApiAuth(CLIENT_ID, CLIENT_SECRET, AUTH_URL).get_auth_token()
    spreadsheet_api = SpreadsheetApi(auth_token)

    multi_document = MultiDocument(CLIENT_ID, CLIENT_SECRET, WORKSPACE_ID, DOCUMENT_ID, SHEET_ID)
    try:
        multi_document.run()
    except requests.HTTPError as e:
        logging.exception(e)

main()
