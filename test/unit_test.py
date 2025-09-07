import unittest
import importlib.util
import sys
from pathlib import Path


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
        CommonTest = load_class_from_file(file_path="../src/CommonTest.py",
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

class TestEnd(unittest.TestCase):
    pass

class TestGetRowNumber(unittest.TestCase):
    pass

class TestFindAllRows(unittest.TestCase):
    pass

class TestFindRowsIntersect(unittest.TestCase):
    pass

class TestFindRowsUnion(unittest.TestCase):
    pass

class TestExpectedValuesCheck(unittest.TestCase):
    pass

class TestSetCellValue(unittest.TestCase):
    pass

class TestAddDataNameResults(unittest.TestCase):
    pass

class TestWriteResults(unittest.TestCase):
    pass

if __name__ == "__main__":
    unittest.main()