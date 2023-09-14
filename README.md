# xlsx教學
## 官方文檔
https://www.npmjs.com/package/xlsx
## 名詞解說
workbook - 可以看做是excel檔案<br/>
worksheet - 即一個excel中的工作表( sheet )<br/>
cell - 指excel中的某一格
## 讀取Excel表格
讀取到的檔案我們稱之為 workbook
```javascript
//1.read方法 ( Extract data from spreadsheet bytes )
//data參數即要讀取的檔案來源，必須是"binary string"
//如 NodeJS buffer or typed array (Uint8Array or ArrayBuffer).
var workbook = XLSX.read(data, opts);

//2.readFile方法( 後端限定，Read spreadsheet bytes from a local file and extract data )
//第一個參數filename要填入檔案路徑
var workbook = XLSX.readFile(filename, opts);

//3備註 : opts 可選參數
//詳細請參考 https://www.npmjs.com/package/xlsx#parsing-options
```
## 處理資料
### 1.建立一個 workbook
一般的電子表格，會需要有至少一個worksheets，<br/>
所以此方法創建出excel後，還需要添加至少一個worksheets，才能輸出成excel檔案
```javascript
//這會創建一個空的workbook( 不含任何worksheets )
var workbook = XLSX.utils.book_new();
```
### 2-1.建立一個 worksheet - 使用二維陣列
下面的aoa是一個二維陣列，裡面裝有 row []
```javascript
var aoa = [  //工作表中的col(直)和row(橫)
  ["A1", "B1", "C1"], //表示工作表A1那格寫上A1...其他依此類推
  ["A2", "B2", "C2"],
  ["A3", "B3", "C3"]
]

var worksheet = XLSX.utils.aoa_to_sheet(aoa, opts);

//opts請參考 https://www.npmjs.com/package/xlsx#array-of-arrays-input
```
### 2-2.建立一個 worksheet - 使用 javascript object
```javascript
var worksheet = XLSX.utils.json_to_sheet(jsa, opts);

// opts 參考 https://www.npmjs.com/package/xlsx#array-of-arrays-input
```
### 3.修改cell的value
sheet_add_aoa方法主要有三個參: worksheet、aoa(二維陣列)、opts<br/>
其中opts最主要使用到的就是origin這個屬性，用來指定要寫入的是哪一個cell
```javascript
//1.修改某一格
XLSX.utils.sheet_add_aoa(worksheet, [[new_value]], { origin: address });
//2.修改多格
XLSX.utils.sheet_add_aoa(worksheet, aoa, opts);

//實際範例:
//以下寫法將會把B3那格寫成1、E5那格寫成abc
XLSX.utils.sheet_add_aoa(worksheet, [
  [1],                             // <-- Write 1 to cell B3
  ,                                // <-- Do nothing in row 4
  [/*B5*/, /*C5*/, /*D5*/, "abc"]  // <-- Write "abc" to cell E5
], { origin: "B3" });
```
### 4.輸出、下載excel檔案
使用write就會自動匯出並下載excel唷!<br/>
預設會輸出的是xlsx File( can be controlled with the bookType property of the opts argument )<br/>
也可以透過更改opts中的type屬性輸出成"binary string", JS string, Uint8Array or Buffer

```javascript

//1.輸出excel檔案
var data = XLSX.write(workbook, opts);

//2.輸出並下載( 保存新檔 )檔案 - 導出多種檔案格式
//所以filename要寫上"檔名+副檔名"，例如"SheetJS.xlsx"
//輸出的附檔就會依據filename中寫的!!
XLSX.writeFile(workbook, filename, opts);

//3.輸出並下載( 保存新檔 )成xlsx檔案
//opts請參考 https://www.npmjs.com/package/xlsx#writing-options
```
## 轉換成JSON或JS Data
### 獲取工作表內容
```javascript
//1.獲取workbook中所有工作表的名稱列表
const sheetNames = workbook.SheetNames  

//2.獲取所有工作表
const sheets = workbook.Sheets

//3.獲取某個工作表
const mySheet = workbook.Sheets[workbook.SheetNames[2]]
```
### 將worksheet轉換成JS 物件
```javascript
//1.轉換成aoa二維陣列(第二維陣列裝的是每roll的cell值)物件
//預設下sheet_to_json方法會從第一行開始掃描，並把該行的值作為標題(header)，所以可以不用設置header
var jsa = XLSX.utils.sheet_to_json(worksheet, opts);

//2.指定header
//下面範例的header:2，指定了要把第2行作為header
var aoa = XLSX.utils.sheet_to_json(worksheet, {...opts, header: 1});
```
以下講解一下，使用 XLSX.utils.sheet_to_json 的範例
```javascript
//假設excel表格內容如下
// | Name  | Age | City    |
// |-------|-----|---------|
// | Alice | 25  | New York|
// | Bob   | 30  | London  |
var aoa = XLSX.utils.sheet_to_json(worksheet, {...opts, header: 1});

//輸出的aoa
// [
//   { "Name": "Alice", "Age": 25, "City": "New York" },
//   { "Name": "Bob", "Age": 30, "City": "London" }
// ]

```


