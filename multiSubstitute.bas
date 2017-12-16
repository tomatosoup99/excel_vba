Attribute VB_Name = "multiSubstitute"
Option Explicit

Function multiSubstitute(str As String, tbl As Range)
'm×2のテーブルを参照してsubstitue関数を適用する
'[使用例]
'str = 'test/YYYY/MM/file.zip'
'tbl =  |YYYY|2017|
        |MM  |12  |
        |DD  |31  |
'結果:'test/2017/12/file.zip'
  
  Dim i As Integer
  For i = 1 To tbl.Row
    str = Application.WorksheetFunction.Substitute(str, tbl(i, 1), tbl(i, 2))
    Debug.Print (tbl(i, 1))
  Next i
  multiSubstitute = str
End Function
