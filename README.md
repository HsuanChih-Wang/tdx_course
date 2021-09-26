# TDX 資料應用課程
##  第一堂課 - 高鐵列車運行圖
修改自張老師提供xlsm檔案裡產生列車運行圖的程式，將大部分vba程式碼抽出來以python改寫。
### 介紹
1. 使用python-excel互動套件xlwings對excel進行讀寫操作。
2. 不幸的是xlwings這個套件並沒有支援在excel繪圖。不過，但他支援從python呼叫excel內置的vba程式。
因此我做了一點修改，把繪圖部分的vba程式保留下來，其餘全數以python改寫。
3. 新增輸出: 可即時掌握當下正在畫哪一車次的圖形。
