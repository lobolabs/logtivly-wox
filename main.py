# -*- coding: utf-8 -*-

from wox import Wox, WoxAPI
from credentials import credentials
import json, datetime, sys, os.path, subprocess

def get_sheet_title_and_column(service, spreadsheetId):
    spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheetId).execute()
    dateCells = "C15:I15"
    for sheet in spreadsheet['sheets']:
        sheetTitle = sheet['properties']['title']
        rangeName = "%s!%s" % (sheetTitle, dateCells)
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheetId, range=rangeName).execute()
        values = result.get('values',[])
        colNum = 0
        # TODO: replace current year with actual cell year, see http://stackoverflow.com/q/42216491/766570
        for row in values:
            for column in row:
                dateStr = "%s %s" % (column, datetime.date.today().year)
                try:
                    cellDate = datetime.datetime.strptime(dateStr, '%b %d %Y')
                    if cellDate.date() == datetime.date.today():
                        return sheetTitle, colNum
                except ValueError:
                    continue

                colNum +=1
        return sheetTitle, colNum

def get_projects_and_hours():
    service, spreadsheetId = credentials.get_service_and_spreadsheetId()
    sheetTitle, colNum = get_sheet_title_and_column(service, spreadsheetId)
    projectCells = 'B16:I19'
    initialProjectCellIndex = 16
    rangeName = "%s!%s" % (sheetTitle, projectCells)
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName, majorDimension='COLUMNS').execute()
    values = result.get('values',[])
    return (
            colNum
            , sheetTitle
            , values and values[0]
            , values and values[colNum+1]
        )

def get_project_cell(projectStr, sheetTitle, colNum):
    service, spreadsheetId = credentials.get_service_and_spreadsheetId()
    cols = ['c','d','e','f','g','h','i']
    columnLetter = cols[colNum]
    # it's not likely that there wil be more than 4 projects at a time
    # but if there is, do logic that fetches all rows before the "total hours" row starts
    projectCells = 'B16:%s19' % columnLetter
    initialProjectCellIndex = 16
    rangeName = "%s!%s" % (sheetTitle, projectCells)
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName, majorDimension='COLUMNS').execute()
    values = result.get('values',[])
    projectNames = list(map((lambda x: x.lower()), values and values[0]))
    rowIndex = [i for i, s in enumerate(projectNames) if projectStr.lower() in s] or [0]
    initialCellValue = values[colNum+1][rowIndex[0]] if len(values) > 1 and rowIndex else 0
    return '%s%s' % (columnLetter, initialProjectCellIndex + rowIndex[0]), float(initialCellValue), projectNames[rowIndex[0]] if projectNames else ''

class Logtively(Wox):
    def AutoComplete(self, project):
        WoxAPI.change_query("log " + project)

    def query(self, query):
        if not os.path.isfile('.credentials\\'+credentials.CREDS_FILENAME):
            subprocess.Popen(['python', 'credentials\credentials.py'], creationflags=8, close_fds=True)
            return [{
                "Title": "Logtively",
                "SubTitle": "Login with your google account and try again",
                "IcoPath": "Images/app.png",
             }]

        colNum, sheetTitle, projects, hours = get_projects_and_hours()

        if query == '':
            index = 0
            results = []
            for project in projects:
                results.append({
                    "Title":project
                    , "SubTitle": "hours logged: "+hours[index]
                    , "IcoPath": "Images/app.png"
                    , "JsonRPCAction":{
                      "method": "AutoComplete"
                      , "parameters":[project]
                      , "dontHideAfterAction":True
                    }
                })
                index+=1
            return results

        args = query.split(' ')
        if len(args) >= 2:
            input_project, input_hours = (' '.join(args[0:-1]), int(args[-1]))
            if type(input_hours) in [int, float] and input_hours not in [0, 0.0]:
                cell, initialCellValue, retrievedProjectStr = get_project_cell(input_project, sheetTitle, colNum)
                rangeName = '%s!%s:%s' % (sheetTitle, cell, cell)
                sys.stderr.write("updating range" + rangeName)
                values = [[initialCellValue + float(input_hours)]]
                body = {
                    'values': values
                }

                service, spreadsheetId =  credentials.get_service_and_spreadsheetId()
                service.spreadsheets().values().update(
                    spreadsheetId=spreadsheetId, range=rangeName, body=body, valueInputOption="USER_ENTERED").execute()
                return [{
                    "Title": "Logtively",
                    "SubTitle": "Added %s hours to %s" % (input_hours, input_project),
                    "IcoPath": "Images/app.png",
                 }]

if __name__ == "__main__":
    Logtively()
