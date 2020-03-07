import os
import sys
import json
import time
import pickle
import gspread
import argparse
from tqdm import tqdm
from pprint import pprint
from collections import defaultdict
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

class SpreadSheetWriter:
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

    def __init__(self, spreadsheet_id, cred_filename='credentials.json'):
        self.sheet = self.__get_sheet()
        self.spreadsheet_id = spreadsheet_id
        self.cred_filename = cred_filename
        self.sheet_id = 0

    def __get_sheet(self):
        creds = None
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    self.cred_filename, self.SCOPES)
                creds = flow.run_local_server(port=0)

            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)

        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()

        return sheet

    def set_sheet_id(self, sheet_id):
        self.sheet_id = sheet_id

    def writerows(self, rows):
        if not rows or len(rows) < 1:
            return

        row_offset = len(rows)
        col_offset = self.__calc_column_offset(rows[0])

        range_name = 'A1:{}{}'.format(col_offset, row_offset)
        body = {
            'values':rows
        }

        value_input_option = 'USER_ENTERED'
        result = self.sheet.values().update(spreadsheetId=self.spreadsheet_id, 
                                            range=range_name, 
                                            valueInputOption=value_input_option, 
                                            body=body).execute()

    def update_backgroundcolor(self, color, row, col):
        if not row or not col:
            return

        requests = []
        requests.append({
            "repeatCell": {
            "cell": {
                "userEnteredFormat": {
                "backgroundColor": {
                    "red": color[0],
                    "green": color[1],
                    "blue": color[2],
                    "alpha": 1
                }}},
            "range": {
                "sheetId": self.sheet_id,
                "startRowIndex": row[0],
                "endRowIndex": row[1],
                "startColumnIndex": col[0],
                "endColumnIndex": col[1]
            },
            "fields": "userEnteredFormat.backgroundColor"
            }}
        )

        body = {
            'requests':requests
        }

        response = self.sheet.batchUpdate(spreadsheetId=self.spreadsheet_id, body=body).execute()

        #print(response)

    def __calc_column_offset(self, row):
        offset = len(row)
        return chr(ord('A') + offset - 1)


class LeetCodeSpreadSheetGenerator:
    def __init__(self, sheetwriter, problems_filename='leetcode.json'):
        self.sheetwriter = sheetwriter
        self.solved_problems = []
        self.problems_info = defaultdict(dict)
        self.problems_filename = problems_filename
        self.color_row_ranges = { 'Easy':[], 'Medium': [], 'Hard': [] }

        self.level_color = {'Easy':   [21/255., 171/255. ,0.], 
                            'Medium': [1.      ,159/255. ,0.], 
                            'Hard':   [1.      ,43/255.  ,0.]}

    def __load(self):
        if not os.path.exists(self.problems_filename):
            return False

        json_obj = None
        with open(self.problems_filename, 'r') as f:
            json_obj = json.load(f)
        
        return json_obj

    def __parse_problems(self):
        json_obj = self.__load()
        if not json_obj:
            return False

        levels = {  1: 'Easy', 
                    2: 'Medium', 
                    3: 'Hard' }

        problems = json_obj['stat_status_pairs']
        for problem in problems:
            stat  = problem['stat']
            level = problem['difficulty']['level']

            question_id    = stat['question_id']
            question_title = stat['question__title']
            question_slug  = stat['question__title_slug']

            self.problems_info[question_id] = { 'title': question_title, 
                                                'slug': question_slug, 
                                                'level':levels[level]}
            
            status = problem['status']
            if status == 'ac':
                self.solved_problems.append(question_id)

        self.solved_problems.sort()

        return True

    def __make_problem_url(self, slug):
        return 'https://leetcode.com/problems/{}/'.format(slug)
    
    def __list_problems(self):
        problems = [['Id', 'Title', 'Difficulty', 'Solved']]
        num_of_problems = len(self.problems_info)
        
        prev_level = None
        level_row_count = 0
        start = 1
        end = 0

        for i, problem_idx in enumerate(range(1, num_of_problems+1)):
            if not self.problems_info.get(problem_idx):
                title = ''
                level = ''
            else:
                slug  = self.problems_info[problem_idx]['slug']
                url = self.__make_problem_url(slug)

                title = '=HYPERLINK(\"{}\",\"{}\")'.format(url, self.problems_info[problem_idx]['title'])
                level = self.problems_info[problem_idx]['level']
                color = self.level_color[level]

            if prev_level != None and prev_level != level:
                if len(prev_level) > 0:
                    self.color_row_ranges[prev_level].append([start, start+level_row_count])

                start = start + level_row_count
                level_row_count = 0
            
            level_row_count += 1
            prev_level = level

            solved = ''
            if i in self.solved_problems:
                solved = u'◯'
            
            problems.append([problem_idx, title, level, solved])

        return problems

    def __update_spreadsheet(self, problems):
        self.sheetwriter.writerows(problems)

        for level, row_range in self.color_row_ranges.items():
            bar = tqdm(row_range)
            bar.set_description(level)
            for start, end in bar:
                self.sheetwriter.update_backgroundcolor(self.level_color[level], [start, end], [2,3])
                time.sleep(1)

    def run(self):
        if not self.sheetwriter:
            return False

        b = self.__parse_problems()
        if not b:
            return False

        problems = self.__list_problems()
        if len(problems) == 0:
            return False

        self.__update_spreadsheet(problems)

        return True

def parse_args():
    parser = argparse.ArgumentParser(prog='gen.py',
                                    description="LeetCode spreadsheet generator")

    parser.add_argument("-i","--spreadsheet_id",
                        required=True,
                        help="Specify Google SpreadSheet ID"
                        )
    parser.add_argument("-f","--credentials_file",
                        default='credentials.json',
                        help="Specify a credentials file of json format"
                        )
    args = parser.parse_args()

    return args

def main(args):
    spreadsheet_id, credentials_filename = args.spreadsheet_id, args.credentials_file
    sheetwriter = SpreadSheetWriter(spreadsheet_id, cred_filename=credentials_filename)
    generator = LeetCodeSpreadSheetGenerator(sheetwriter)
    generator.run()
    
if __name__ == '__main__' :
    args = parse_args()
    main(args)

