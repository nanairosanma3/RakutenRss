import openpyxl as pyxl
import time
import subprocess

#####参考サイト：https://zenn.dev/ririkabu/articles/b7adbbd3012eea


class OrderBookMonitor:
    def __init__(self, excel_path, code_list, json_base_path):
        self.excel_path     = excel_path
        self.code_list      = code_list
#        self.json_base_path = Path(json_base_path)
        self.pif            = None # Exel を起動するprocess id
        self.previous_hashes = {}
        self.headers        = []
#        self.data_queue     = MPQueue # Multiprocessing Queue
        self.stop_flag      = False
        self.reader_process = None # データ取得プロセスの参照

        # `data_range` を code_list の範囲から動的に設定
        start_row = 2
        end_row   = start_row + len( code_list ) - 1
#        self.data_range = f" B{ start_now } : ES{ end_row } "




    # excelファイルを新規作成し、アクティブなシートのA1から
    # 市況情報関数のヘッダー行を関数ごとに表示
    def create_excel(self):
        wb = pyxl.Workbook()
        ws = wb.active
        ws["A1"].value = "=RssMarketHeader(1)" # 1 = 国内株式、2 = 先物、3 = 指数、4 = 為替

        #市場データを取得するセルの設定
        # row = 行, col = column = 列
        # for i in range( A, B ) --> i はAから( B - 1 )まで変化する
        for row in range(2, 502):
            for col in range(2, 150):
                cell = ws.cell( row = row, column = col )
                # RssMaket("銘柄コード", "取得項目"), pyxl.utils.get_column_letter = 数字をアルファベットに
                cell.value = f"= RssMarket( $A{row}, { pyxl.utils.get_column_letter(col) } $1 )"

        # 銘柄コードの設定
        for i, code in enumerate(self.code_list):
            ws[ f"A{ i + 2 }" ].value = str(code)


        wb.save(self.excel_path)

    def get_path_to_xl2(self) -> Path:
        try:
            subprocess_rtn = subprocess.run(["assoc",".xlsx"], shell=True, stdout=subprocess.PIPE, stderr=subproess.PIPE)
            assoc_to = re.search(r"Excel\.Sheet\.\d+", subprocess_rtn.stdout.deode("utf-8")).group()

            subprocess_rtn = subprocess.run(["ftype", assoc_to], shell=True, stdout=subprocess.PIPE, stderr=Subprocess.PIPE)
            xl_path = re.search(r"C:.*EXCEL\.EXE", subprocess_rtn.stdut.decode("uyf-8")).group

            print("Excelのパス:", xl_path) # デバック用出力
            return Path(xl_path)

        except (AttributeError, subprocess.CalledProcessError) as e:
            raise FileNotFoundError("Excelのパスが取得できませんでした。Excelがインストールされているか確認してください。") from e






# def add_xl_app(self) -> xw.App:
#    try:
        # エクセル実行プログラム"C:.*EXCEL\.EXE"を探して返す
#        excel_path = self.get_path_to_xl2()
        # /x オプションを使用してアドインを有効化
#        command = f"'{ str(excel_path )}' /x '{ self.excel_path }'"

#        proc = subprocess.Popen(command)
#        time.sleep(1)   # 初期化待機

        # プロセスの確認と接続, 0 から 10 まで
#        for _ in range(10)
#            try:
#                xl_app = xw.apps[proc.pid]
#                print("PID:", proc.pid, "アプリケーションが正常に起動しました。")
#                return xl_App, proc.pid
#            except KeyError:
#                print("PID確認中...")
#                time.sleep(1)
        
            # /x オプションでアドインを有効にした状態でExcelを起動
            # subprocess.Popen 非同期でExcelプロセスを起動
            # PID管理　起動したExcelプロセスを追跡、xlwingsで使う








#####参考ここまで



