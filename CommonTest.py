import openpyxl
import pandas as pd
import glob

class CommonTest:
    def __init__(self):
        self.total_rows = []
        self.total_cols = []
        self.workbook   = None
        self.results_ws = None
        self.active_ws  = None
        self.currentResultsRow = None
        
        self._rowCol    = 1
        self._nameCol   = 2
        self._expValCol = 3
        self._operCol   = 4
        self._actValCol = 5
        self._checkCol  = 6

    def initializeTest(self, csv_files, output_file):
        with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
            self.workbook = writer.book
            header_format = self.workbook.add_format({'bold': True})
            for index, csv_file in enumerate(csv_files):
                sheet_name = f"AnalyzedData-{index}"
                df = pd.read_csv(csv_file)
                df.to_excel(writer, sheet_name=sheet_name, index=False)   
                self.active_ws = writer.sheets[sheet_name]

                for col_num, value in enumerate(df.columns.values):
                    self.active_ws.writer(0, col_num, value, header_format)
                
                for i, col_num  in enumerate(df.columns):
                    max_length = max(df[col_num].astype(str).map(len).map(), len(col_num))
                    self.active_ws.set_column(i, i, max_length + 2)

                self.active_ws.freeze_panes(1,0)
                rows, cols = df.shape
                
                self.total_rows.append(rows)
                self.total_cols.append(cols)
        
        self._createResultsFile()
        self.workbook.close()
        self._openResultsWorkbook(output_file)

    def endTest(self, output_file):
        self.workbook.save(output_file) 

    def _createResultsFile(self):
        self.results_ws = self.workbook.add("Results")
        self._formatResultsWorsheet()

    def _formatResultsWorsheet(self):
        header_format = self.workbook.add_format({
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "border": 1
        })

        headers = ["Row_Index","Name","Actual Value",
                    "Operation","Expected Value"]

        for col_num, header in enumerate(headers):
            self.results_ws.write(0, col_num, header, header_format)

        self.results_ws.freeze_panes(1,0)

    def _openResultsWorkbook(self, output_file):
        self.workbook   = openpyxl.load_workbook(output_file)
        self.active_ws  = self.workbook.active
        self.results_ws = self.workbook["Results"]

    def _activateWorksheet(self, csvSheet):
        self.active_ws = self.workbook[f"AnalyzedData-{csvSheet}"]
    
    def getRowNumber(self, searchString, colNum, csvSheet):
        self._activateWorksheet(csvSheet)
        iterator = self.active_ws.iter_rows(min_col=colNum, 
                                            max_col=colNum,
                                            values_only=True)
        
        for rowNum, rowVal in enumerate(iterator, start=self.active_ws.min_row):
            if rowVal[0] == searchString:
                return rowNum

    def findAllRows(self, searchString, colNum, csvSheet):   
        allRows = []
        self._activateWorksheet(csvSheet)
        iterator = self.active_ws.iter_rows(min_col=colNum, 
                                            max_col=colNum,
                                            values_only=True)
        
        for rowNum, rowVal in enumerate(iterator, start=self.active_ws.min_row):
            if rowVal[0] == searchString:
                allRows.append(rowNum)
        
        return allRows

    def findRowsIntersect(self, searchStringDict, csvSheet):
        intersection = None
        for key, value in searchStringDict.items():
            allRows = set(self.findAllRows(searchString=value, 
                                           colNum=key, 
                                           csvSheet=csvSheet))
            
            intersection =  allRows if intersection is None else (intersection & allRows)

        return list(intersection)

    def findRowsUnion(self, searchStringDict, csvSheet):
        union = set()
        for key, value in searchStringDict.items():
            union = list(set(union) | set(self.findAllRows(searchString=value, 
                                                           colNum=key, 
                                                           csvSheet=csvSheet)))

        return list(union)
    
    def _returnStrCoordRC(self):
        actVal = f'RC[{self._actValCol - self._checkCol}]'
        expVal = f'RC[{self._expValCol - self._checkCol}]'
        return actVal, expVal
    
    def _commandEquals(self):
        actVal, expVal = self._returnStrCoordRC()
        return f"=IF({actVal} = {expVal}, {self._passVal}, {self._failVal})"

    def _commandNotEquals(self):
        actVal, expVal = self._returnStrCoordRC()
        return f"=IF({actVal} = {expVal}, {self._passVal}, {self._failVal})"

    def _commandWithtinTolerance(self, tol):
        actVal, expVal = self._returnStrCoordRC()
        check1 = f"{actVal} >= {expVal} + {tol}"
        check2 = f"{actVal} <= {expVal} - {tol}"
        formula = f"AND({check1}, {check2})"
        return f"=IF({formula}, {self._passVal}, {self._failVal})"

    def _commandOutsideTolerance(self, tol):
        actVal, expVal = self._returnStrCoordRC()
        check1 = f"{actVal} >= {expVal} + {tol}"
        check2 = f"{actVal} <= {expVal} - {tol}"
        formula = f"NOT(AND({check1}, {check2}))"
        return f"=IF({formula}, {self._passVal}, {self._failVal})"

    def expectedValuesCheck(self, expectedValue, actualValue):
        stringSplit = expectedValue.trim().split(",", 1)
        commandStr = stringSplit[0]

        if commandStr in ("EQ","SEQ"):
            self._commandEquals()
        elif commandStr in ("NE","SNE"):
            self._commandNotEquals()
        elif commandStr in ("TL"):
            lastComma = stringSplit[1].rfind(",")
            self._commandWithtinTolerance(tol=stringSplit[1][lastComma+1:])
        elif commandStr in ("NTL"):
            lastComma = stringSplit[1].rfind(",")
            self._commandOutsideTolerance(tol=stringSplit[1][lastComma+1:])
        else:
            # unsuported command
            pass

    def _increaseResultsRow(self):   
        self.currentResultsRow += 1
    
    def addDataNameResults(self, titleStr, dataRow, colNum, cmmt):
        pass

    def wrtieResults(self, titleStr, dataRow, expectedValue, 
                     actualValue, colNum=None, cmmt=None):
        
        self.addDataNameResults(titleStr=titleStr, 
                                dataRow=dataRow, 
                                colNum=colNum, 
                                cmmt=cmmt)
        
        self.expectedValuesCheck(expectedValue=expectedValue,
                                 actualValue=actualValue)

        self._increaseResultsRow()