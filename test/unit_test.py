import unittest
import importlib.util
import sys
import time
import os
import pandas as pd
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor

def load_class_from_file(file_path, class_name):

    file_path = Path(file_path).resolve()
    module_name = file_path.stem  # filename without .py

    # Load the module dynamically
    spec = importlib.util.spec_from_file_location(module_name, file_path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)

    # Get the class object
    cls = getattr(module, class_name)
    return cls 

class TestInitialize(unittest.TestCase):
    def setUp(self):
        self.file_dir = str(Path(__file__).resolve().parent)
        self.csv_files = [
            "/../csv_data/realistic_data_1.csv",
            "/../csv_data/realistic_data_2.csv",
            "/../csv_data/realistic_data_3.csv",
            "/../csv_data/realistic_data_4.csv",
            "/../csv_data/realistic_data_5.csv",
            "/../csv_data/realistic_data_6.csv",
            "/../csv_data/realistic_data_7.csv",
            "/../csv_data/realistic_data_8.csv",
            "/../csv_data/realistic_data_9.csv",
            "/../csv_data/realistic_data_10.csv"
        ]
        self.output_file = "/../results/realistic_data.xlsx"
        CommonTest = load_class_from_file(file_path=self.file_dir + "/../src/CommonTest.py",
                                          class_name="CommonTest")
        self.cls = CommonTest()
    
    def tearDown(self):
        if Path(self.file_dir + self.output_file).exists():
            Path(self.file_dir + self.output_file).unlink()

    def test_createFile(self):
        self.cls.initializeTest(
            csv_files=[self.file_dir + self.csv_files[0]],
            output_file=self.file_dir + self.output_file
        )

        self.assertEqual(Path(self.file_dir + self.output_file).exists(), True)
        self.assertEqual(len(self.cls.workbook.sheetnames), 2)
        self.assertEqual("Results" in self.cls.workbook.sheetnames, True)
        self.assertEqual(f"AnalyzedData-{1}" in self.cls.workbook.sheetnames, True)

    def test_createFileMultCSVs(self):
        self.cls.initializeTest(
            csv_files=list(map(lambda x: self.file_dir + x, self.csv_files)),
            output_file=self.file_dir + self.output_file
        )

        self.assertEqual(Path(self.file_dir + self.output_file).exists(), True)
        self.assertEqual(len(self.cls.workbook.sheetnames), 11)
        self.assertEqual("Results" in self.cls.workbook.sheetnames, True)
        
        for i in range(len(self.csv_files)):
            self.assertEqual(f"AnalyzedData-{i+1}" in self.cls.workbook.sheetnames, True)

    def test_dataImport(self):
        csvTest = list(map(lambda x: self.file_dir + x, self.csv_files))
        self.cls.initializeTest(
            csv_files=csvTest,
            output_file=self.file_dir + self.output_file
        )

        for index, paths in enumerate(csvTest):
            df_csv = pd.read_csv(paths, dtype=str)
            df_xlsx = pd.read_excel(self.file_dir + self.output_file,
                                   sheet_name=f"AnalyzedData-{index+1}",
                                   dtype=str)
            
            self.assertEqual(df_csv.equals(df_xlsx), True)
        
class TestEnd(unittest.TestCase):
    def setUp(self):
        self.file_dir = str(Path(__file__).resolve().parent)
        self.csv_files = [
            "/../csv_data/realistic_data_1.csv",
            "/../csv_data/realistic_data_2.csv",
            "/../csv_data/realistic_data_3.csv",
            "/../csv_data/realistic_data_4.csv",
            "/../csv_data/realistic_data_5.csv",
            "/../csv_data/realistic_data_6.csv",
            "/../csv_data/realistic_data_7.csv",
            "/../csv_data/realistic_data_8.csv",
            "/../csv_data/realistic_data_9.csv",
            "/../csv_data/realistic_data_10.csv"
        ]
        self.output_file = "/../results/realistic_data.xlsx"
        CommonTest = load_class_from_file(file_path=self.file_dir + "/../src/CommonTest.py",
                                          class_name="CommonTest")
        self.cls = CommonTest()
    
    def tearDown(self):
        if Path(self.file_dir + self.output_file).exists():
            Path(self.file_dir + self.output_file).unlink()

    def test_saveFile(self):
        self.cls.initializeTest(
            csv_files=[self.file_dir + self.csv_files[0]],
            output_file=self.file_dir + self.output_file
        )

        if Path(self.file_dir + self.output_file).exists():
            Path(self.file_dir + self.output_file).unlink()

        self.assertEqual(Path(self.file_dir + self.output_file).exists(), False)
        self.cls.endTest(self.file_dir + self.output_file)
        self.assertEqual(Path(self.file_dir + self.output_file).exists(), True)

class TestGetCellInfo(unittest.TestCase):
    def setUp(self):
        self.file_dir = str(Path(__file__).resolve().parent)
        self.csv_files = [
            "/../csv_data/realistic_data_1.csv",
            "/../csv_data/realistic_data_2.csv",
            "/../csv_data/realistic_data_3.csv",
            "/../csv_data/realistic_data_4.csv",
            "/../csv_data/realistic_data_5.csv",
            "/../csv_data/realistic_data_6.csv",
            "/../csv_data/realistic_data_7.csv",
            "/../csv_data/realistic_data_8.csv",
            "/../csv_data/realistic_data_9.csv",
            "/../csv_data/realistic_data_10.csv"
        ]
        self.output_file = "/../results/realistic_data.xlsx"
        CommonTest = load_class_from_file(file_path=self.file_dir + "/../src/CommonTest.py",
                                          class_name="CommonTest")
        self.cls = CommonTest()

        self.cls.initializeTest(
            csv_files=list(map(lambda x: self.file_dir + x, self.csv_files)),
            output_file=self.file_dir + self.output_file
        )

    def tearDown(self):
        if Path(self.file_dir + self.output_file).exists():
            Path(self.file_dir + self.output_file).unlink()

    def test_getRowNum(self):
        self.assertEqual(self.cls.getRowNumber("jennifer39@yahoo.com", 3), 503)
        self.assertEqual(self.cls.getRowNumber("Ruth", 1), 14)
        self.assertEqual(self.cls.getRowNumber("6477", self.cls.total_cols[4] - 1, 5), 1001)
        self.assertEqual(self.cls.getRowNumber(6477, self.cls.total_cols[4] - 1, 5), None)

    def test_getColNum(self):
        self.assertEqual(self.cls.getColumnNumber("CreditCard"), 13)
        self.assertEqual(self.cls.getColumnNumber("DateOfBirth", 10), 12)
        self.assertEqual(self.cls.getColumnNumber("Testing", 10), None)

    def test_findAllRows(self):
        self.assertEqual(len(self.cls.findAllRows("David",1,1)), 16)
        self.assertEqual(len(self.cls.findAllRows("Missouri",7,10)), 27)
        self.assertEqual(len(self.cls.findAllRows("Hello World",7,10)), 0)

    def test_findRowsIntersect(self):
        strDict = {
            1: "David",
            7: "Florida"
        }

        listRows = self.cls.findRowsIntersect(strDict,10)
        self.assertEqual(len(listRows), 1)
        self.assertEqual(listRows[0], 29)

        listRows = self.cls.findRowsIntersect(strDict,1)
        self.assertEqual(len(listRows), 0)

    def test_findRowsUnion(self):
        strDict = {
            1: "David",
            7: "Florida"
        }

        rowsName = self.cls.findAllRows(strDict[1], 1, 10)
        rowsState = self.cls.findAllRows(strDict[7], 7, 10)
        listRows = self.cls.findRowsUnion(strDict,10)
        self.assertEqual(len(listRows), len(rowsName) + len(rowsState) - 1)

        rowsName = self.cls.findAllRows(strDict[1], 1, 1)
        rowsState = self.cls.findAllRows(strDict[7], 7, 1)
        listRows = self.cls.findRowsUnion(strDict,1)
        self.assertEqual(len(listRows), len(rowsName) + len(rowsState))

        rowsName = self.cls.findAllRows("dfsjdf", 1, 1)
        rowsState = self.cls.findAllRows("sdkfmdskf", 7, 1)
        listRows = self.cls.findRowsUnion({1:"dfsjdf",2:"sdkfmdskf"},1)
        self.assertEqual(len(listRows), len(rowsName) + len(rowsState))

    def test_getCellValue(self):
        self.assertEqual(self.cls.getCellValue(2,1), "Jessica")
        self.assertEqual(self.cls.getCellValue(500,1, 10), "Anthony")
        self.assertEqual(self.cls.getCellValue(5,4), "(170)522-9895")
        self.assertEqual(self.cls.getCellValue(10000,4, 10), None)

class TestExpectedValuesCheck(unittest.TestCase):
    pass

class TestSetCellValue(unittest.TestCase):
    pass

class TestAddDataNameResults(unittest.TestCase):
    pass

class TestWriteResults(unittest.TestCase):
    pass

class TestParellelProcess(unittest.TestCase):
    
    time_sequential = 0
    time_parallel = 0

    def setUp(self):
        self.file_dir = str(Path(__file__).resolve().parent)
        self.csv_files = [
            "/../csv_data/realistic_data_1.csv",
            "/../csv_data/realistic_data_2.csv",
            "/../csv_data/realistic_data_3.csv",
            "/../csv_data/realistic_data_4.csv",
            "/../csv_data/realistic_data_5.csv",
            "/../csv_data/realistic_data_6.csv",
            "/../csv_data/realistic_data_7.csv",
            "/../csv_data/realistic_data_8.csv",
            "/../csv_data/realistic_data_9.csv",
            "/../csv_data/realistic_data_10.csv"
        ]

    def tearDown(self):
        self.removeFiles()
    
    def removeFiles(self):
        with os.scandir(self.file_dir + "/../results/") as entries:
            for entry in entries:
                if entry.is_file() and entry.name.lower().endswith(".xlsx"):
                    os.remove(entry.path)

    def helperFunc(self, output_file):
        CommonTest = load_class_from_file(file_path=self.file_dir + "/../src/CommonTest.py",
                                          class_name="CommonTest")
        cls = CommonTest()
        csvTest = list(map(lambda x: self.file_dir + x, self.csv_files))
        cls.initializeTest(
            csv_files=csvTest,
            output_file=self.file_dir + output_file
        )
        
        for index, paths in enumerate(csvTest):
            df_csv = pd.read_csv(paths, dtype=str)
            df_xlsx = pd.read_excel(self.file_dir + output_file,
                                   sheet_name=f"AnalyzedData-{index+1}",
                                   dtype=str)
            
            self.assertEqual(df_csv.equals(df_xlsx), True)
       
    def sequentialFunc(self, num_files=3):
        start_seq = time.time()
        
        for i in range(num_files):
             self.helperFunc(output_file=f"/../results/realistic_data_{i}.xlsx")
        
        end_seq = time.time()
        return end_seq - start_seq
    
    def parellelFunc(self, num_files=3):
        start_mp = time.time()
        files = [f"/../results/realistic_data_{i}.xlsx" for i in range(num_files)]
        with ThreadPoolExecutor(max_workers=3) as executor:
            executor.map(self.helperFunc, files)

        end_mp = time.time()
        return end_mp - start_mp

    def test_xcompare(self):
        time_seq = self.sequentialFunc()
        self.removeFiles()
        time_thread = self.parellelFunc()
        self.removeFiles()
        self.assertLess(time_seq, time_thread)


if __name__ == "__main__":
    unittest.main()