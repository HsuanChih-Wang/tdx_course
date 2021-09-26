import xlwings as xw
import random
# import 兩種package: xlwings 和 random

def getStationY(stationName):  # 以車站名稱比對車站Y軸座標 並回傳Y軸
    statioinList = {"南港": 90, "台北": 120, "板橋": 150, "桃園": 180, "新竹": 210, "苗栗": 240,
                    "台中": 270, "彰化": 300, "雲林": 330, "嘉義": 360, "台南": 390, "左營": 420}
    y = statioinList[stationName]  # 以list儲存並讀取
    return y


def getEndRow(x):  # GetEndRow 為找相同車次最後迄站之列數
    sheet = workbook.sheets['HSRSchedule']  # 車次資料在 工作表 HSRScheule
    while sheet.cells(x + 1, 8).value != '1' and sheet.cells(x + 1, 8).value != None:  # '取下一筆資料之站序非 1
        x = x + 1
    endRow = x
    return endRow  # 回傳該車次最後車次 之列數


# drawline 畫出運行線
def draw_line(Red, Green, Blue, i, type):
    # 接受引數: 紅, 綠, 藍三種顏色隨機數, 列數i, 型態type
    sheet = workbook.sheets['HSRSchedule']  # 車次資料在 工作表 HSRScheule

    def getHour(str):
        # 取得小時
        hour = int(str[0:2])  # 字串切割後轉型
        return hour

    def getMinute(str):
        # 取得分鐘
        min = int(str[3:])  # 字串切割後轉型
        return min

    if type == 0:
        # 以現在站名與發車時間及下一站站名與到站時間為做 起迄座標 畫出運行線
        # GetStationY 為以站名 比對各站的 Y座標 , X座標則以發車時間轉成分鐘
        bx = int(getHour(sheet.cells(i, 12).value) * 60 + getMinute(sheet.cells(i, 12).value))
        by = getStationY(sheet.cells(i, 10).value)
        ex = int(getHour(sheet.cells(i + 1, 11).value) * 60 + getMinute(sheet.cells(i + 1, 11).value))
        ey = getStationY(sheet.cells(i + 1, 10).value)
    else:  # type = 1
        # 以下一站站名與到站時間及下一站站名與發車時間為做 座標 畫出在下一站 列車停留時間運行線
        # GetStationY 為以站名 比對各站的 Y座標 , X座標則以車次時間轉成分鐘
        bx = int(getHour(sheet.cells(i + 1, 11).value) * 60 + getMinute(sheet.cells(i + 1, 11).value))
        by = getStationY(sheet.cells(i + 1, 10).value)
        ex = int(getHour(sheet.cells(i + 1, 12).value) * 60 + getMinute(sheet.cells(i + 1, 12).value))
        ey = getStationY(sheet.cells(i + 1, 10).value)

    draw_line_func = workbook.macro('Moudle.Drawline')  # 呼叫存在excel裡面的vba函數 Drawline 進行畫圖
    draw_line_func(Red, Green, Blue, bx, by, ex, ey)  # 傳入需要的引數


if __name__ == '__main__':

    # 讓 xlwings 動態開啟excel檔案，並將該檔案存入 workbook 變數
    workbook = xw.Book(r'學員_列車運行圖_修改.xlsm')
    # 開啓 '運行圖' 試算表
    sheet = workbook.sheets['運行圖']

    ## 以下開始與張老師課程提供vba程式基本上一致 ##
    sheet.range('A1:EO14').column_width = 4.78  #設定欄位寬度
    sheet.range('A1:EO14').row_height = 30  #設定列高
    sheet.range('B3:EO25').delete(shift='left')  #清除原始圖並移到左側

    sheet = workbook.sheets['HSRSchedule']  #設定要讀取的試算表: HSRSchedule
    row = 2  #從班表第一車開始畫

    while sheet.cells(row, 1).value != None:  # '當 Sheets("HSRSchedule") 有資料時.執行While
        # 注意: 原本vba 寫 <> "" python需改成 != None　否則進入無窮迴圈!
        red = int(255 * random.random())  # 取亂數顏色為每一次車次運行線之顏色
        green = int(255 * random.random())
        blue = int(255 * random.random())
        endRow = getEndRow(row)  # 找相同車次最後迄站之列數

        trainNo = sheet.cells(row, 2).value  #車次編號
        startStation = sheet.cells(row, 5).value  #起站
        endStation = sheet.cells(row, 7).value  #迄站
        print("開始畫→ 車次: {0} 起站: {1} 迄站: {2}" .format(trainNo, startStation, endStation))

        for i in range(row, endRow):
            sheet = workbook.sheets['HSRSchedule']  # 回工作表 HSRSchedule
            draw_line(red, green, blue, i, 0)
            draw_line(red, green, blue, i, 1)

        row = endRow + 1  #從endRow的下一列繼續

    print("全部車次已畫完！")





