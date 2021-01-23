"""
   前回のユーザー入力を記憶するクラス

   wxglade_out.pyに依存している
   AccTempGuiクラスに入れても良かったんだけど、見通しが悪くなるので別にした

"""
import os
import pickle


class LastUserInput:
    """
    前回のユーザー入力を記憶するクラス
    """

    def __init__(self, gui):
        """
        コンストラクタ
        :param gui: GUIオブジェクト
        """
        self.gui = gui
        self.file_name = None

    def make_file_name(self):
        """
        記憶するファイルのファイル名を生成する（入力ファイル名の拡張子を.actpに変えたもの）
        :return:
        """
        filename = self.gui.text_ctrl_input_filename.GetValue()
        # 拡張子だけ変更する
        pre, ext = os.path.splitext(filename)
        self.file_name = pre + ".actp"
        print(self.file_name)

    def load_user_input(self):
        """
        ファイルからユーザー入力を読み出す
        :return:
        """
        self.make_file_name()
        try:
            with open(self.file_name, mode="rb") as f:
                self.restore_user_input(f)
                return True

        except FileNotFoundError as e:
            print(e)
            return False

        except Exception as e:
            self.gui.error_process(e)
            return False

    def restore_user_input(self, f):
        """
        ファイルからユーザー情報を読み出し、各GUIコントロールにセットする
        :param f: 読み出すファイル
        :return:
        """
        load_data = pickle.load(f)  # Dict型のload_data変数にファイルの内容を読み出す
        print(load_data)

        self.gui.text_ctrl_input_filename.SetValue(load_data["text_ctrl_input_filename"])  # 入力ファイル名（冗長だけど、、、）
        self.gui.combo_box_prefecture.SetStringSelection(load_data["combo_box_prefecture"])  # 県名
        self.gui.onPrefectureSelected(None)  # 県名が確定したらイベントを発生させ地点名リストを得る
        self.gui.combo_box_station.SetStringSelection(load_data["combo_box_station"])  # 地点名をセットする
        self.gui.onStationSelected(None)  # 地点名が確定したらイベントを発生させ気象庁クラスに地点名を通知する
        self.gui.spin_ctrl_title_line_no.SetValue(load_data["spin_ctrl_title_line_no"])  # 行番号をセットする
        self.gui.onStationSelected(None)  # 行番号が確定したイベント
        self.gui.check_list_box_target_sheet_name.SetCheckedStrings(
            load_data["check_list_box_target_sheet_name"])  # シート名
        self.gui.onTargetSheetSelected(None)  # 行番号とシート名が確定したらカラム名を取得する為イベントを発生させる
        self.gui.combo_box_start_date.SetStringSelection(load_data["combo_box_start_date"])
        self.gui.combo_box_start_date.SetStringSelection(load_data["combo_box_start_date"])
        self.gui.combo_box_end_date.SetStringSelection(load_data["combo_box_end_date"])
        self.gui.combo_box_target_temperature.SetStringSelection(load_data["combo_box_target_temperature"])
        self.gui.combo_box_current_temperature.SetStringSelection(load_data["combo_box_current_temperature"])
        self.gui.combo_box_rate.SetStringSelection(load_data["combo_box_rate"])
        self.gui.combo_box_estimate_date.SetStringSelection(load_data["combo_box_estimate_date"])
        self.gui.text_ctrl_output_filename.SetValue(load_data["text_ctrl_output_filename"])

    def save_user_input(self):
        """
        ユーザー入力を保存する
        :return:
        """
        try:
            with open(self.file_name, mode="wb") as f:
                self.store_user_input(f)
                return True

        except Exception as e:
            self.gui.error_process(e)
            return False

    def store_user_input(self, f):
        """
        ユーザーが入力したデータをファイルに出力する
        :param f: 出力ファイル
        :return:
        """
        # 辞書型の変数として格納
        save_data = {"text_ctrl_input_filename": self.gui.text_ctrl_input_filename.GetValue(),
                     "combo_box_prefecture": self.gui.combo_box_prefecture.GetStringSelection(),
                     "combo_box_station": self.gui.combo_box_station.GetStringSelection(),
                     "spin_ctrl_title_line_no": self.gui.spin_ctrl_title_line_no.GetValue(),
                     "check_list_box_target_sheet_name": self.gui.check_list_box_target_sheet_name.GetCheckedStrings(),
                     "combo_box_start_date": self.gui.combo_box_start_date.GetStringSelection(),
                     "combo_box_end_date": self.gui.combo_box_end_date.GetStringSelection(),
                     "combo_box_target_temperature": self.gui.combo_box_target_temperature.GetStringSelection(),
                     "combo_box_current_temperature": self.gui.combo_box_current_temperature.GetStringSelection(),
                     "combo_box_rate": self.gui.combo_box_rate.GetStringSelection(),
                     "combo_box_estimate_date": self.gui.combo_box_estimate_date.GetStringSelection(),
                     "text_ctrl_output_filename": self.gui.text_ctrl_output_filename.GetValue(),
                     }
        pickle.dump(save_data, f)  # 作った変数をファイルに保存
