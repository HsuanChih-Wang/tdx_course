# TDX 資料應用課程
##  高鐵列車運行圖
如果你和我一樣對於vba語法不熟悉，習慣用python做數據處理，那xlwings會是好幫手!
修改自張老師提供xlsm檔案裡產生列車運行圖(學員_列車運行圖.xlsm)的程式，將大部分vba程式碼抽出來以python改寫。

### 介紹
1. 使用python-excel互動套件xlwings對excel進行讀寫操作。
2. 原本打算全部都改成python，但不幸的是xlwings這個套件並沒有支援在excel繪圖。
不過，他支援從python呼叫excel內置的vba程式。因此我做了一點修改: 把繪圖部分的vba程式保留下來，其餘全數以python改寫。
3. 新增輸出: 可即時掌握當下正在畫哪一車次的圖形。
4. 執行預覽: 
![image](https://user-images.githubusercontent.com/53686476/134799305-b53476e1-0cc8-47ca-baea-941aabd6d96f.png)


### 使用方法
1. 自行clone或選取Download ZIP
![image](https://user-images.githubusercontent.com/53686476/134798900-ec3a91d1-0622-48ed-ad5c-22ba0af25930.png)
2. 執行HSR_Schedule.py

* 執行前需先安裝xlwings套件
