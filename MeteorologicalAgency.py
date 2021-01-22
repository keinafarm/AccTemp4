# -*- coding: utf-8 -*-

"""
気象庁から平均気温と平年気温のデータを取得する

https://gist.github.com/barusan/3f098cc74b92fad00b9bb4478da35385
を参照した

"""
from datetime import date
import urllib.request
import lxml.html
from datetime import datetime as dt


def encode_data(data):
    """
    Map型のオブジェクトをパーセントエンコードされた ASCII 文字列に変換する
    :param data: 変換するmap型オブジェクト
    :return: 変換したURL文字列
    """
    return urllib.parse.urlencode(data).encode(encoding='ascii')


def get_phpsessid():
    """
    セッションIDを得る
    気象庁とやりとりする為のPHPSESSIDを取得する
    :return:PHPSESSID
    """

    URL = "http://www.data.jma.go.jp/gmd/risk/obsdl/index.php"  # セッションIDEAを取得する為のURL
    xml = urllib.request.urlopen(URL).read().decode("utf-8")
    tree = lxml.html.fromstring(xml)
    return tree.cssselect("input#sid")[0].value


def get_station_by_id(pd):
    """
    指定したIDのSTATION ID Listを得る
    :param pd: 0:全国 0以外:station id listを得る県のid
    :return: station id list
    """

    def kansoku_items(bits):
        """
        取得する観測データ：気温のみ
        :param bits:
        :return:
        """
        return dict(rain=(bits[0] == "1"),
                    wind=(bits[1] == "1"),
                    temp=(bits[2] == "1"),
                    sun=(bits[3] == "1"),
                    snow=(bits[4] == "1"))

    def parse_station(dom):
        stitle = dom.get("title").replace("：", ":")
        title = dict(filter(lambda y: len(y) == 2,
                            map(lambda x: x.split(":"), stitle.split("\n"))))

        name = title["地点名"]
        stid = dom.cssselect("input[name=stid]")[0].value
        stname = dom.cssselect("input[name=stname]")[0].value
        kansoku = kansoku_items(dom.cssselect("input[name=kansoku]")[0].value)
        assert name == stname
        return stname, dict(id=stid, flags=kansoku)

    def parse_prefs(dom):
        name = dom.text
        prid = int(dom.cssselect("input[name=prid]")[0].value)
        return name, prid

    URL = "http://www.data.jma.go.jp/gmd/risk/obsdl/top/station"
    data = encode_data({"pd": "%02d" % pd})
    xml = urllib.request.urlopen(URL, data=data).read().decode("utf-8")
    tree = lxml.html.fromstring(xml)

    if pd > 0:
        station_ids = dict(map(parse_station, tree.cssselect("div.station")))  # 地区名とstation IDの対応辞書を作成
    else:
        station_ids = dict(map(parse_prefs, tree.cssselect("div.prefecture")))  # 県名とprefecture IDの対応辞書を作成

    return station_ids


def download_temperature_csv(phpsessid, station, element, start_date, end_date):
    """
    気温データをcsv形式でダウンロードする
    :param phpsessid: セッションID
    :param station: 地点 ID
    :param element: 取得データID
    :param start_date: 開始日付
    :param end_date: 終了日付
    :return: 取得したCSVデータ
    """
    params = {  # サイトに送るForm Data
        "PHPSESSID": phpsessid,
        # 共通フラグ
        "rmkFlag": 1,  # 利用上注意が必要なデータを格納する
        "disconnectFlag": 1,  # 観測環境の変化にかかわらずデータを格納する
        "csvFlag": 1,  # すべて数値で格納する
        "ymdLiteral": 1,  # 日付は日付リテラルで格納する
        "youbiFlag": 0,  # 日付に曜日を表示する
        "kijiFlag": 0,  # 最高・最低（最大・最小）値の発生時刻を表示
        # 日別値データ選択
        "aggrgPeriod": 1,  # 日別値
        "stationNumList": '["%s"]' % station,  # 観測地点IDのリスト
        "elementNumList": '[["%s",""]]' % element,  # 項目IDのリスト
        "ymdList": '["%d", "%d", "%d", "%d", "%d", "%d"]' % (
            start_date.year, end_date.year,
            start_date.month, end_date.month,
            start_date.day, end_date.day),  # 取得する期間
        "jikantaiFlag": 0,  # 特定の時間帯のみ表示する
        "interAnnualFlag": 1,  # 連続した期間で表示する
        'jikantaiList': [1, 24],
        "optionNumList": ' [["op1", 0]]',  # 平年値も得る [["op1",0]]
        "downloadFlag": "true",  # CSV としてダウンロードする？
        "huukouFlag": 0,
    }

    print("load")
    URL = "http://www.data.jma.go.jp/gmd/risk/obsdl/show/table"
    data = encode_data(params)
    csv = urllib.request.urlopen(URL, data=data).read().decode("cp932", "ignore")
    # なぜか、文字コードの変換エラーが出る事があるが、どうせタイトル行は削るので、無視する
    #    csv = urllib.request.urlopen(URL, data=data).read().decode("shift-jis")
    print("complete")
    return csv


class PrefectureList:
    """
    県単位のリスト
    """

    def __init__(self):
        """
        県リストを取得する
        """
        self.data = get_station_by_id(0)

    def get_list(self):
        """
        県名の文字列リストを得る
        :return:県名の文字列リスト
        """
        string_list = self.data.keys()
        return list(string_list)

    def get_id(self, prefecture):
        """
        指定した県名に対するIDを得る
        :param prefecture:
        :return:県名に対するID
        """
        return self.data[prefecture]


class StationList:
    def __init__(self, prefecture_id):
        """
        地点リストを取得する
        """
        self.prefecture_id = prefecture_id
        self.data = get_station_by_id(prefecture_id)
        print(self.data)
        # 気温をサポートしていない地点を省く
        temp_dict = self.data.copy()
        for item in temp_dict:
            if not self.data[item]['flags']['temp']:
                del self.data[item]

    def get_list(self):
        """
        地点の文字列リストを得る
        :return:県名の文字列リスト
        """
        string_list = self.data.keys()
        return list(string_list)

    def get_id(self, station):
        """
        指定した県名に対するIDを得る
        :param station: 地点名
        :return:地点に対するID
        """
        station_data = self.data[station]
        return station_data['id']


class MeteorologicalAgency:
    """
    気象庁のサイトからデータを取得するクラス
    """

    def __init__(self):
        """
        取得する地点を指定する
        """
        self.prefecture = None
        self.station = None
        self.csv = None
        self.prefecture_obj = None
        self.prefecture_id = None
        self.station_obj = None
        self.station_id = None

    def get_prefecture_list(self):
        self.prefecture_obj = PrefectureList()
        prefecture_list = self.prefecture_obj.get_list()
        return prefecture_list

    def set_prefecture(self, prefecture):
        self.prefecture = prefecture
        self.prefecture_id = self.prefecture_obj.get_id(prefecture)
        self.station_obj = StationList(self.prefecture_id)

    def get_station_list(self):
        return self.station_obj.get_list()

    def set_station(self, station):
        self.station = station
        self.station_id = self.station_obj.get_id(station)

    def get_temperature_string(self, start_date, end_data):
        """
        指定した期間の平均気温と平年値のリストをCSV形式で得る
        :param start_date: 期間の開始日
        :param end_data: 期間の終了日
        :return: 気象庁から送られてきた気温データの文字列（CSV)
        """

        phpsess_id = get_phpsessid()
        self.csv = download_temperature_csv(phpsess_id, self.station_id, 201,
                                            start_date, end_data)
        return self.csv

    def get_temperature_list(self, start_date, end_data):
        """
        指定した期間の平均気温と平年気温のリストを得る
        :param start_date: 期間の開始日
        :param end_data: 期間の終了日
        :return: [[日付(date型),平均気温(float),平年気温(float)]...]のリスト
        """
        self.get_temperature_string(start_date, end_data)
        out_list = []
        for item in self.csv.splitlines(True)[6:]:  # 最初の6行はタイトルなので無視する
            line_data = item.split(',')  # カンマで分割する
            if line_data[4] == '':
                line_data[4] = line_data[1]  # 平年値が含まれていない場合は去年のデータにする
            data_pair = [dt.strptime(line_data[0], '%Y/%m/%d'), float(line_data[1]), float(line_data[4])]
            # 日付型と浮動小数に変換する
            out_list.append(data_pair)
        return out_list


if __name__ == "__main__":
    obj = MeteorologicalAgency()
    print(obj.get_prefecture_list())
    obj.set_prefecture("高知")
    #    obj.set_prefecture("根室")
    print(obj.get_station_list())
    obj.set_station("窪川")
    #    obj.set_station("羅臼")
    temperature_list = obj.get_temperature_list(date(2019, 12, 26), date(2020, 12, 25))

    print(temperature_list)
