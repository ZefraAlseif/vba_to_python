import sys
import importlib.util
from pathlib import Path
import os
import pandas as pd

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

time_sequential = 0
time_parallel = 0

file_dir = str(Path(__file__).resolve().parent)
csv_files = [
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
output_file = "/../results/realistic_data_2.xlsx"

def removeFiles():
    os.remove(file_dir + output_file)

def helperFunc():
    CommonTest = load_class_from_file(file_path=file_dir + "/../src/CommonTest.py",
                                        class_name="CommonTest")
    cls = CommonTest()
    csvTest = list(map(lambda x: file_dir + x, csv_files))
    cls.initializeTest(
        csv_files=csvTest,
        output_file= file_dir + output_file
    )
    
    for index, paths in enumerate(csvTest):
        df_csv = pd.read_csv(paths, dtype=str)
        df_xlsx = pd.read_excel(file_dir + output_file,
                                sheet_name=f"AnalyzedData-{index+1}",
                                dtype=str)

if __name__ == "__main__":
    helperFunc()
    removeFiles()