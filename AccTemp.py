"""
   積算温度をエクセルっファイルに出力するアプリケーション
"""
__version__ = "0.9.0"
__author__ = "Akira Yoshida <akiracraftwork@infoseek.jp>"

#   このクラスが、メインのクラスになります

from wxglade_out import AccTempFrame
from MeteorologicalAgency import MeteorologicalAgency
from ATWorkBook import ATWorkBook
from Calculation import Calculation
from LastUserInput import LastUserInput
import openpyxl
import wx
import traceback


class AccTempGui(AccTempFrame):
    """
    メインクラス
        wxGladeで作成されたAccTempFrameクラスを継承(wxglade_out.py)し、
        wxglade_out.pyは直接編集しなくても良いようにする
    """

    def __init__(self, *args, **kwargs):
        """
        画面の初期状態を作成する
        :param args:
        :param kwargs:
        """
        super(AccTempGui, self).__init__(*args, **kwargs)
        self.meteorological_agency = MeteorologicalAgency()
        prefecture_list = self.meteorological_agency.get_prefecture_list()
        self.combo_box_prefecture.SetItems(prefecture_list)
        self.workbook = None
        self.calculation = None
        self.last_user_input = LastUserInput(self)
        title = self.GetTitle()
        title += "  Version: " + __version__
        self.SetTitle(title)

    def onInputFileEnter(self, event):  # wxGlade: AccTempFrame.<event_handler>
        """
        入力ファイルが指定された時の処理
        :param event:
        :return:
        """
        filename = self.text_ctrl_input_filename.GetValue()  # 入力ファイル名を取得
        try:
            self.workbook = ATWorkBook(filename)  # EXCELファイルとして読み込む
            self.workbook.target_sheet_all()  # シート名一覧を得るためにすべてのシートを処理対象にする
            sheet_name_list = self.workbook.sheets_name_list  # シート名リストを得る
            self.check_list_box_target_sheet_name.SetItems(sheet_name_list)  # シート名選択コントロールにシート名リストをセット

            text = filename + "を読み込みました"
            self.text_ctrl_message.SetValue(text)
            self.last_user_input.load_user_input()  # 前回のユーザー入力を復元する

        except openpyxl.utils.exceptions.InvalidFileException as e:  # 不正なファイルが指定された場合
            error_text = "指定された'" + filename + "'は、EXCELファイルではありません"
            self.text_ctrl_message.SetValue(error_text)
            print(e)
            print(type(e))  # ログ用
            return

        except FileNotFoundError as e:  # 指定されたファイルが存在しない場合
            error_text = "指定された'" + filename + "'は、存在しません"
            self.text_ctrl_message.SetValue(error_text)  # メッセージ表示欄に表示
            print(e)
            print(type(e))  # ログ用
            return

        except Exception as e:  # 予期していないエラー
            self.error_process(e)  # 予期していないエラーの処理
            return

    def OnChooseInputFile(self, event):  # wxGlade: AccTempFrame.<event_handler>
        """
        入力ファイルボタンが押された
        :param event:
        :return:
        """
        style = wx.FD_DEFAULT_STYLE | wx.FD_OPEN | wx.FD_FILE_MUST_EXIST | wx.FD_CHANGE_DIR
        # 読み込み用のファイル選択ダイアログのモード指定
        pathname = self.showFileDialog(style)  # 読み込み用のファイル選択ダイアログを表示
        if pathname is None:
            return  # 何も選択されていない
        self.text_ctrl_input_filename.SetValue(pathname)  # ファイル名を取得
        print("input file is " + pathname)  # ログ用
        self.onInputFileEnter(event)  # 入力ファイルが指定された時の処理

    def OnChooseOutputFile(self, event):  # wxGlade: AccTempFrame.<event_handler>
        """
        出力ファイルボタンが押された
        :param event:
        :return:
        """
        style = wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT | wx.FD_CHANGE_DIR
        # 書き込み用のファイル選択ダイアログのモード指定
        pathname = self.showFileDialog(style)  # 書き込み用のファイル選択ダイアログを表示
        if pathname is None:
            return  # 何も指定されていない
        self.text_ctrl_output_filename.SetValue(pathname)  # 指定されたファイル名を表示
        print("output file is " + pathname)  # ログ用

    def onPrefectureSelected(self, event):  # wxGlade: AccTempFrame.<event_handler>
        """
        県が選択された
        :param event:
        :return:
        """
        selected = self.combo_box_prefecture.GetStringSelection()  # 選択された県名を得る
        print("onPrefectureSelected:" + selected)
        self.meteorological_agency.set_prefecture(selected)  # 気象庁オブジェクトに通知
        selection_list = self.meteorological_agency.get_station_list()  # 地点リストを得る
        self.combo_box_station.SetItems(selection_list)  # 地点の選択リストに地点リストを設定
        if self.combo_box_station.GetStringSelection() == "":  # デバッグ時の確認用
            print("Station:指定されてない")
        else:
            print("Station:[[" + self.combo_box_station.GetStringSelection() + "]]")

    def onStationSelected(self, event):  # wxGlade: AccTempFrame.<event_handler>
        """
        地点が選択された
        :param event:
        :return:
        """
        selected = self.combo_box_station.GetStringSelection()  # 選択された地点名を得る
        print("onStationSelected:" + selected)
        self.meteorological_agency.set_station(selected)  # 気象庁オブジェクトに通知

    def onTargetSheetSelected(self, event):  # wxGlade: AccTempFrame.<event_handler>
        """
        対象となるシートリストが選択された
        :param event:
        :return:
        """
        self.set_column_name_list()  # カラム名の選択リストをコンボボックスにセットする

    def onColumnNameSelected(self, event):  # wxGlade: AccTempFrame.<event_handler>
        """
        各カラム名が選択された時の処理

        最初は、選択されるたびに、他のコンボボックスから候補を減らしていこうと思ったのだけど
        それやり始めたら、とても面倒だという気持ちになったので、ここでは何もしない。
        :param event:
        :return:
        """
        selected_combo_box = event.EventObject  # 変更されたコンボボックスを得る
        column_name = selected_combo_box.GetStringSelection()  # 指定された値を得る
        print(column_name)

    def onTitleLineNoChanged(self, event):  # wxGlade: AccTempFrame.<event_handler>
        """
        カラム名の行番号が変更された
        :param event:
        :return:
        """
        self.set_column_name_list()  # カラム名の選択リストをコンボボックスにセットする

    def OnChooseExit(self, event):  # wxGlade: AccTempFrame.<event_handler>
        """
        終了ボタンが押された
        :param event:
        :return:
        """
        self.Destroy()  # アプリから抜ける

    def showFileDialog(self, style):
        """
        ファイル選択ダイアログを表示する
        :param style: wx.FileDialogへのパラメータ https://wxpython.org/Phoenix/docs/html/wx.FileDialog.html
        :return: 指定されたファイルパス名(キャンセル時はNone)
        """
        with wx.FileDialog(self, 'ファイルを選択してください', style=style) as dialog:
            if dialog.ShowModal() == wx.ID_CANCEL:
                return None  # キャンセルボタンが押されたらNoneで返る
            return dialog.GetPath()  # 指定されたファイルパス名を返す

    def set_prefecture_list(self, prefecture_list):
        """
        県名リストを表示する
        :param prefecture_list: 県名リスト
        :return:
        """
        self.combo_box_prefecture.SetItems(prefecture_list)

    def set_column_name_list(self):
        """
        カラム名を指定するコンボボックスのリストを設定
        :return:
        """
        if self.workbook is None:  # 入力ファイルが指定されていない時（無いはず）
            print("入力Fileが指定されていない")
            error_text = "入力ファイルが指定されていません"
            self.text_ctrl_message.SetValue(error_text)
            return
        sheet_name_list = self.check_list_box_target_sheet_name.GetCheckedStrings()
        # 読み込んだファイルのシート一覧を得る
        if len(sheet_name_list) == 0:
            print("No Sheet Name")
            return  # シート名が入力されていない
        line_no = self.spin_ctrl_title_line_no.GetValue()  # タイトル行の行番号を得る
        if line_no <= 0:
            print("No Line No.")
            return  # 行番号が入力されていない
        self.workbook.set_target_sheets_by_name_list(sheet_name_list)  # 対象となるシートの一覧を指定する
        print(sheet_name_list)
        print(self.workbook.sheets_list)

        # カラム名の一覧を得る
        column_data_list = self.workbook.get_row_data_string(line_no)  # 指定したシート上の指定した行番号の各セルをテキストデータとして得る
        column_data = set()  # 名前の重複を省くためset型変数に格納する
        for row in column_data_list:  # 取得したテキストデータ分
            column_data |= set(row)  # 重複しないように、和集合でセットする
        column_data = list(sorted(column_data))  # あとあとsetだと使いにくいので、ソートシタリストに変換
        print(column_data)

        # それぞれのコンボボックスに設定する
        self.combo_box_start_date.SetItems(column_data)  # 開始日は必須
        column_data.insert(0, "--使用しない--")
        self.combo_box_end_date.SetItems(column_data)  # それ以外は使用しない事が選択可能
        self.combo_box_target_temperature.SetItems(column_data)
        self.combo_box_current_temperature.SetItems(column_data)
        self.combo_box_rate.SetItems(column_data)
        self.combo_box_estimate_date.SetItems(column_data)

        self.combo_box_end_date.SetSelection(0)  # 初期状態として "--使用しない--"を表示
        self.combo_box_target_temperature.SetSelection(0)
        self.combo_box_current_temperature.SetSelection(0)
        self.combo_box_rate.SetSelection(0)
        self.combo_box_estimate_date.SetSelection(0)

    def check_input_data(self):
        """
        入力データをチェック
        :return: True=妥当 False=誤り
        """
        error_text = ''  # 誤りにした要因を格納する変数
        if self.workbook is None:
            error_text += " 入力ファイルが指定されていません\n"
        if self.combo_box_prefecture.GetStringSelection() == "":
            error_text += " 県が指定されていません\n"
        if self.combo_box_station.GetStringSelection() == "":
            error_text += " 地点が指定されていません\n"
        if self.spin_ctrl_title_line_no.GetValue() <= 0:
            error_text += " タイトルの行番号が指定されていません\n        if len(self.check_list_box_target_sheet_name.GetCheckedStrings()) == 0:
            error_text += " 対象となるシート名が指定されていません\n"
        if self.combo_box_start_date.GetStringSelection() == '':
            error_text += " 積算温度起算日のカラムが指定されていません\n"
        if self.text_ctrl_output_filename.GetValue() == "":
            error_text += " 出力ファイル名が指定されていません\n"

        if len(error_text) > 0:  # いずれかの誤りがあった
            error_text = "入力内容が正しくありません\n" + error_text
            self.text_ctrl_message.SetValue(error_text)  # エラーメッセージを表示
            return False
        else:  # 入力データは問題ない
            self.text_ctrl_message.SetValue("")  # メッセージ欄をクリア
            print("Start")
            return True

    def get_combo_string(self, combo_box):
        """
        コンボボックスの内容を得る
        :param combo_box: 取得するコンボボックス
        :return: コンボボックス値（"--使用しない--"が選択されている時はNone）
        """
        combo_string = combo_box.GetStringSelection()
        if combo_string == "--使用しない--":
            combo_string = None
        return combo_string

    def get_column_names(self):
        """
        ユーザーが入力した、各カラム名を得る
        :return:
        """
        self.workbook.set_target_work_sheet(self.check_list_box_target_sheet_name.GetCheckedStrings())
        # 対象とするシートを、ユーザーが指定したシートにする

        line_no = self.spin_ctrl_title_line_no.GetValue()  # タイトル行番号を得る
        start_date = self.combo_box_start_date.GetStringSelection()  # 開始日を取得するカラム名を得る

        end_date = self.get_combo_string(self.combo_box_end_date)  # 終了日を取得するカラム名を得る
        target_temperature = self.get_combo_string(self.combo_box_target_temperature)  # 目標積算温度を取得するカラム名を得る
        current_temperature = self.get_combo_string(self.combo_box_current_temperature)  # 現状積算温度を格納するカラム名を得る
        rate = self.get_combo_string(self.combo_box_rate)  # 到達率を格納するカラム名を得る
        estimate_date = self.get_combo_string(self.combo_box_estimate_date)  # 予測終了日を格納するカラム名を得る

        self.workbook.set_column_position_string(line_no, start_date, end_date, target_temperature, current_temperature,
                                                 rate, estimate_date)  # それぞれのカラム名を通知する

    def message(self, text):
        """
        メッセージ表示欄にテキストを出力する
        :param text: 出力するテキスト
        :return:
        """
        self.text_ctrl_message.SetValue(text)

    def OnChooseRun(self, event):  # wxGlade: AccTempFrame.<event_handler>
        """
        計算して保存ボタンが押された
        :param event:
        :return:
        """
        if not self.check_input_data():  # 入力されたデータをチェックする
            return
        self.get_column_names()  # 各カラム名をワークシートオブジェクトに通知する

        try:
            self.calculation = Calculation(self.workbook, self.meteorological_agency, self)  # 積算温度計算オブジェクトを作る
            self.calculation.run()  # 積算温度計算実行
            self.workbook.save(self.text_ctrl_output_filename.GetValue())  # 結果を保存する
            self.text_ctrl_message.SetValue(self.text_ctrl_output_filename.GetValue() + "に保存しました")
            self.last_user_input.save_user_input()  # ユーザーが入力したデータを記憶する
        except PermissionError as e:  # 書き込み禁止エラー
            error_text = "指定された'" + self.text_ctrl_output_filename.GetValue() + "'に、書き込めません。\n他のアプリケーションで開いている可能性があります。"
            self.text_ctrl_message.SetValue(error_text)
            print(e)
            print(type(e))
            return

        except Exception as e:  # 予期しないエラー
            self.error_process(e)
            return

    def error_process(self, e):
        """
        予期しない例外が発生した時の処理
        :param e:
        :return:
        """
        print(e)
        print(type(e))
        print(traceback.format_exc())

        error_text = "Error: type:{0} : {1}\n".format(type(e), e) + traceback.format_exc()
        self.text_ctrl_message.SetValue(error_text)


class AccTempApp(wx.App):
    """
    wxPythonお決まりのアプリケーションクラス
    """

    def OnInit(self):
        self.frame = AccTempGui(None, wx.ID_ANY, "")
        self.SetTopWindow(self.frame)
        self.frame.Show()
        return True


# end of class AccTempApp

if __name__ == "__main__":
    app = AccTempApp(0)
    app.MainLoop()
