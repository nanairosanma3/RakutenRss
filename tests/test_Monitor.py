# 「標準ライブラリ  サードパーティライブラリ ローカルで作成したモジュール」の順に一行ずつあけて書く
# 上から読み込まれるので上書きを防ぐ
import pytest
import os
import time
import sys
import subprocess
from dotenv import load_dotenv
sys.path.append(os.getcwd())
from getStockValue import OrderBookMonitor



load_dotenv()
EXCEL_PATH     = TEST_EXCEL_PATH = os.getenv("TEST_EXCEL_PATH")
code_list      = ["1234", "5678", "9012"]
json_base_path = TEST_JSON_PATH  = os.getenv("TEST_JSON_PATH")


Test = OrderBookMonitor(TEST_EXCEL_PATH, code_list, TEST_JSON_PATH)

def test_create_excel():
    Test.create_excel()
    assert os.path.isfile(TEST_EXCEL_PATH)


