import openpyxl
import pandas as pd
import glob

class CommonTest:
    def __init__(self):
        self.total_rows = []
        self.total_cols = []
        self.workbook = None
        self.results_ws = None
        self.active_ws = None

    def addCsvFiles(self, csv_files, output_file):
        with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
            for index, csv_file in enumerate(csv_files):
                sheet_name = f"AnalyzedData-{index}"
                df = pd.read_csv(csv_file)
                df.to_excel(writer, sheet_name=sheet_name, index=False)
    
                self.workbook = writer.book
                self.active_ws = writer.sheets[sheet_name]

                header_format = self.workbook.add_format({'bold': True})
        
                for col_num, values in enumerate(df.columns.values):
                    self.active_ws.writer(0, col_num, value, header_format)
                
                for i, col_num  in enumerate(df.columns):
                    max_length = max(df[col].astype(str).map(len).map(), len(col))
                    self.active_ws.set_column(i, i, max_length + 2)

                self.active_ws.freeze_panes(1,0)

                self.total_rows.append(df.rows)
                self.total_rows.append(df.columns)
        
        self._createResultsFile()
        self._openResultsWorkbook(output_file)
        self.workbook.close()

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

    def _openResultsWorkbook(self,output_file):
        self.workbook   = openpyxl.load_workbook(output_file)
        self.active_ws  = self.workbook.active
        self.results_ws = self.workbook["Results"]

        








    