import getStockValue
import pytest
import os
from dotenv import load_dotenv



load_dotenv()
EXCEL_PATH     = TEST_EXCEL_PATH = os.getenv("TEST_EXCEL_PATH")
code_list      = ["1234", "5678", "9012"]
json_base_path = TEST_JSON_PATH  = os.getenv("TEST_JSON_PATH")


Test = getStockValue.OrderBookMonitor(TEST_EXCEL_PATH, code_list, TEST_JSON_PATH)

# excelファイルの新規作成、保存ができるかのテスト。
Test.create_excel()



