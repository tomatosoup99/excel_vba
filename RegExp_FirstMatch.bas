Attribute VB_Name = "RegExp_FirstMatch"
Option Explicit

Function RegExp_FirstMatch(ptn As String, str As String)
'正規表現で最初にマッチした文字列を返す
'参考サイト:https://msdn.microsoft.com/ja-jp/library/cc392403.aspx
'[使用例]
'ptn = b(a|b|c)
'str = aabbcc
'結果: bb
  
  Dim regEx, Matches
  Set regEx = CreateObject("VBScript.RegExp")
  regEx.Pattern = ptn
  regEx.IgnoreCase = False
  regEx.Global = True
  Set Matches = regEx.Execute(str)
    
  RegExp_FirstMatch = Matches(0)
End Function

