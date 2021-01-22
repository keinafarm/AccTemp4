#!/usr/bin/env python3
# -*- coding: UTF-8 -*-
#
# generated by wxGlade 1.0.1 on Fri Jan 22 18:17:15 2021
#

import wx

# begin wxGlade: dependencies
# end wxGlade

# begin wxGlade: extracode
# end wxGlade


class AccTempFrame(wx.Frame):
    def __init__(self, *args, **kwds):
        # begin wxGlade: AccTempFrame.__init__
        kwds["style"] = kwds.get("style", 0) | wx.DEFAULT_FRAME_STYLE
        wx.Frame.__init__(self, *args, **kwds)
        self.SetSize((614, 656))
        self.SetTitle(u"積算温度をEXCELファイルに書き込む")

        self.panel_1 = wx.Panel(self, wx.ID_ANY)

        sizer_1 = wx.BoxSizer(wx.VERTICAL)

        sizer_2 = wx.BoxSizer(wx.VERTICAL)
        sizer_1.Add(sizer_2, 0, wx.EXPAND, 0)

        sizer_3 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_2.Add(sizer_3, 0, wx.EXPAND, 0)

        label_1 = wx.StaticText(self.panel_1, wx.ID_ANY, u"入力ファイル")
        sizer_3.Add(label_1, 0, wx.ALIGN_CENTER_VERTICAL, 0)

        self.text_ctrl_input_filename = wx.TextCtrl(self.panel_1, wx.ID_ANY, "", style=wx.TE_PROCESS_ENTER | wx.TE_READONLY)
        sizer_3.Add(self.text_ctrl_input_filename, 2, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)

        self.button_input_file = wx.Button(self.panel_1, wx.ID_ANY, "...")
        sizer_3.Add(self.button_input_file, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)

        sizer_5 = wx.BoxSizer(wx.VERTICAL)
        sizer_2.Add(sizer_5, 0, wx.EXPAND, 0)

        sizer_6 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_5.Add(sizer_6, 0, wx.EXPAND, 0)

        label_3 = wx.StaticText(self.panel_1, wx.ID_ANY, u"県")
        sizer_6.Add(label_3, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)

        self.combo_box_prefecture = wx.ComboBox(self.panel_1, wx.ID_ANY, choices=[], style=wx.CB_DROPDOWN | wx.CB_READONLY)
        sizer_6.Add(self.combo_box_prefecture, 0, wx.ALL, 5)

        label_4 = wx.StaticText(self.panel_1, wx.ID_ANY, u"地点")
        sizer_6.Add(label_4, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)

        self.combo_box_station = wx.ComboBox(self.panel_1, wx.ID_ANY, choices=[], style=wx.CB_DROPDOWN | wx.CB_READONLY)
        sizer_6.Add(self.combo_box_station, 0, wx.ALL, 5)

        label_6 = wx.StaticText(self.panel_1, wx.ID_ANY, u"タイトルの行番号")
        sizer_6.Add(label_6, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)

        self.spin_ctrl_title_line_no = wx.SpinCtrl(self.panel_1, wx.ID_ANY, "", min=0, max=65535)
        sizer_6.Add(self.spin_ctrl_title_line_no, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)

        sizer_7 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_2.Add(sizer_7, 1, wx.EXPAND, 0)

        grid_sizer_1 = wx.FlexGridSizer(7, 2, 1, 1)
        sizer_7.Add(grid_sizer_1, 0, 0, 0)

        label_5 = wx.StaticText(self.panel_1, wx.ID_ANY, u"対象となるシート名")
        grid_sizer_1.Add(label_5, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)

        self.check_list_box_target_sheet_name = wx.CheckListBox(self.panel_1, wx.ID_ANY, choices=[], style=wx.LB_MULTIPLE)
        self.check_list_box_target_sheet_name.SetMinSize((99, 80))
        grid_sizer_1.Add(self.check_list_box_target_sheet_name, 2, wx.ALL, 5)

        label_7 = wx.StaticText(self.panel_1, wx.ID_ANY, u"積算温度起算日のカラム")
        grid_sizer_1.Add(label_7, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)

        self.combo_box_start_date = wx.ComboBox(self.panel_1, wx.ID_ANY, choices=[], style=wx.CB_DROPDOWN | wx.CB_READONLY)
        self.combo_box_start_date.SetMinSize((150, 23))
        grid_sizer_1.Add(self.combo_box_start_date, 2, wx.ALL, 5)

        label_8 = wx.StaticText(self.panel_1, wx.ID_ANY, u"積算終了日のカラム")
        grid_sizer_1.Add(label_8, 0, wx.ALL, 5)

        self.combo_box_end_date = wx.ComboBox(self.panel_1, wx.ID_ANY, choices=[], style=wx.CB_DROPDOWN | wx.CB_READONLY)
        self.combo_box_end_date.SetMinSize((150, 23))
        grid_sizer_1.Add(self.combo_box_end_date, 2, wx.ALL, 5)

        label_9 = wx.StaticText(self.panel_1, wx.ID_ANY, u"積算温度目標値のカラム")
        grid_sizer_1.Add(label_9, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)

        self.combo_box_target_temperature = wx.ComboBox(self.panel_1, wx.ID_ANY, choices=[], style=wx.CB_DROPDOWN | wx.CB_READONLY)
        self.combo_box_target_temperature.SetMinSize((150, 23))
        grid_sizer_1.Add(self.combo_box_target_temperature, 2, wx.ALL, 5)

        label_10 = wx.StaticText(self.panel_1, wx.ID_ANY, u"算出した積算温度のカラム")
        grid_sizer_1.Add(label_10, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)

        self.combo_box_current_temperature = wx.ComboBox(self.panel_1, wx.ID_ANY, choices=[], style=wx.CB_DROPDOWN | wx.CB_READONLY)
        self.combo_box_current_temperature.SetMinSize((150, 23))
        grid_sizer_1.Add(self.combo_box_current_temperature, 2, wx.ALL, 5)

        label_11 = wx.StaticText(self.panel_1, wx.ID_ANY, u"目標値に対する到達率のカラム")
        grid_sizer_1.Add(label_11, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)

        self.combo_box_rate = wx.ComboBox(self.panel_1, wx.ID_ANY, choices=[], style=wx.CB_DROPDOWN | wx.CB_READONLY)
        self.combo_box_rate.SetMinSize((150, 23))
        grid_sizer_1.Add(self.combo_box_rate, 2, wx.ALL, 5)

        label_12 = wx.StaticText(self.panel_1, wx.ID_ANY, u"目標到達予測日のカラム")
        grid_sizer_1.Add(label_12, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)

        self.combo_box_estimate_date = wx.ComboBox(self.panel_1, wx.ID_ANY, choices=[], style=wx.CB_DROPDOWN | wx.CB_READONLY)
        self.combo_box_estimate_date.SetMinSize((150, 23))
        grid_sizer_1.Add(self.combo_box_estimate_date, 2, wx.ALL, 5)

        label_13 = wx.StaticText(self.panel_1, wx.ID_ANY, u"【使い方】\n① 入力ファイルの右側のボタンを押して読み込むExcelファイルを指定します。\n② 観測地点の県を選びます\n③ 観測地点を選びます\n④ タイトルの名前がある行番号を選びます\n⑤ 積算温度を計算するシートを選びます\n⑥ 積算温度の起点となる日付が格納されているタイトルの名前を選びます\n⑦ 残りのカラムのタイトルの名前を選びます\n⑧ 出力するファイル名を指定します。\n⑨ 「計算して保存」ボタンを押すと入力ファイルを読み込み、気象庁から気温データを読み込み、積算温度を計算して、出力ファイルに書き込みます。\n\n次回に入力ファイルが指定された時に、前回の入力した各パラメータを思い出します\n")
        sizer_7.Add(label_13, 2, wx.ALL | wx.EXPAND, 5)

        sizer_4 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_2.Add(sizer_4, 0, wx.EXPAND, 0)

        label_2 = wx.StaticText(self.panel_1, wx.ID_ANY, u"出力ファイル")
        sizer_4.Add(label_2, 0, wx.ALIGN_CENTER_VERTICAL | wx.RESERVE_SPACE_EVEN_IF_HIDDEN, 0)

        self.text_ctrl_output_filename = wx.TextCtrl(self.panel_1, wx.ID_ANY, "", style=wx.TE_READONLY)
        sizer_4.Add(self.text_ctrl_output_filename, 2, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)

        self.button_output_file = wx.Button(self.panel_1, wx.ID_ANY, "...")
        sizer_4.Add(self.button_output_file, 0, wx.ALIGN_CENTER_VERTICAL, 0)

        sizer_8 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_2.Add(sizer_8, 1, wx.EXPAND, 0)

        sizer_8.Add((100, 20), 2, wx.ALIGN_CENTER_VERTICAL, 0)

        self.button_run = wx.Button(self.panel_1, wx.ID_ANY, u"計算して保存")
        sizer_8.Add(self.button_run, 1, wx.ALL, 5)

        sizer_8.Add((100, 20), 2, wx.ALIGN_CENTER_VERTICAL, 0)

        self.button_2 = wx.Button(self.panel_1, wx.ID_ANY, u"終了")
        sizer_8.Add(self.button_2, 1, wx.ALL, 5)

        sizer_8.Add((100, 20), 2, wx.ALIGN_CENTER_VERTICAL, 0)

        self.text_ctrl_message = wx.TextCtrl(self.panel_1, wx.ID_ANY, "", style=wx.TE_MULTILINE | wx.TE_READONLY)
        sizer_2.Add(self.text_ctrl_message, 100, wx.ALL | wx.EXPAND, 5)

        grid_sizer_1.AddGrowableCol(1)

        self.panel_1.SetSizer(sizer_1)

        self.Layout()

        self.Bind(wx.EVT_TEXT_ENTER, self.onInputFileEnter, self.text_ctrl_input_filename)
        self.Bind(wx.EVT_BUTTON, self.OnChooseInputFile, self.button_input_file)
        self.Bind(wx.EVT_COMBOBOX, self.onPrefectureSelected, self.combo_box_prefecture)
        self.Bind(wx.EVT_COMBOBOX, self.onStationSelected, self.combo_box_station)
        self.Bind(wx.EVT_SPINCTRL, self.onTitleLineNoChanged, self.spin_ctrl_title_line_no)
        self.Bind(wx.EVT_CHECKLISTBOX, self.onTargetSheetSelected, self.check_list_box_target_sheet_name)
        self.Bind(wx.EVT_COMBOBOX, self.onColumnNameSelected, self.combo_box_start_date)
        self.Bind(wx.EVT_COMBOBOX, self.onColumnNameSelected, self.combo_box_end_date)
        self.Bind(wx.EVT_COMBOBOX, self.onColumnNameSelected, self.combo_box_target_temperature)
        self.Bind(wx.EVT_COMBOBOX, self.onColumnNameSelected, self.combo_box_current_temperature)
        self.Bind(wx.EVT_COMBOBOX, self.onColumnNameSelected, self.combo_box_rate)
        self.Bind(wx.EVT_COMBOBOX, self.onColumnNameSelected, self.combo_box_estimate_date)
        self.Bind(wx.EVT_BUTTON, self.OnChooseOutputFile, self.button_output_file)
        self.Bind(wx.EVT_BUTTON, self.OnChooseRun, self.button_run)
        self.Bind(wx.EVT_BUTTON, self.OnChooseExit, self.button_2)
        # end wxGlade

    def onInputFileEnter(self, event):  # wxGlade: AccTempFrame.<event_handler>
        print("Event handler 'onInputFileEnter' not implemented!")
        event.Skip()

    def OnChooseInputFile(self, event):  # wxGlade: AccTempFrame.<event_handler>
        print("Event handler 'OnChooseInputFile' not implemented!")
        event.Skip()

    def onPrefectureSelected(self, event):  # wxGlade: AccTempFrame.<event_handler>
        print("Event handler 'onPrefectureSelected' not implemented!")
        event.Skip()

    def onStationSelected(self, event):  # wxGlade: AccTempFrame.<event_handler>
        print("Event handler 'onStationSelected' not implemented!")
        event.Skip()

    def onTitleLineNoChanged(self, event):  # wxGlade: AccTempFrame.<event_handler>
        print("Event handler 'onTitleLineNoChanged' not implemented!")
        event.Skip()

    def onTargetSheetSelected(self, event):  # wxGlade: AccTempFrame.<event_handler>
        print("Event handler 'onTargetSheetSelected' not implemented!")
        event.Skip()

    def onColumnNameSelected(self, event):  # wxGlade: AccTempFrame.<event_handler>
        print("Event handler 'onColumnNameSelected' not implemented!")
        event.Skip()

    def OnChooseOutputFile(self, event):  # wxGlade: AccTempFrame.<event_handler>
        print("Event handler 'OnChooseOutputFile' not implemented!")
        event.Skip()

    def OnChooseRun(self, event):  # wxGlade: AccTempFrame.<event_handler>
        print("Event handler 'OnChooseRun' not implemented!")
        event.Skip()

    def OnChooseExit(self, event):  # wxGlade: AccTempFrame.<event_handler>
        print("Event handler 'OnChooseExit' not implemented!")
        event.Skip()

# end of class AccTempFrame

class AccTempApp(wx.App):
    def OnInit(self):
        self.frame = AccTempFrame(None, wx.ID_ANY, "")
        self.SetTopWindow(self.frame)
        self.frame.Show()
        return True

# end of class AccTempApp

if __name__ == "__main__":
    app = AccTempApp(0)
    app.MainLoop()