Option Explicit 

'実行例：Cscript ShowFileProperty.vbs ディレクトリ名
' 実行例：Cscript ShowFileProperty.vbs ディレクトリ名 >D:\FilePropertyList.txt

Dim objArgs 
Set objArgs = WScript.Arguments 

If objArgs.count = 1 Then
 GET_DETAILS(objArgs.item(0)) 
Else
 Wscript.Echo "引数が不正です。" & vbCrLf & "実行例：Cscript ShowFileProperty.vbs ディレクトリ名"
 Wscript.Quit '終了 
End If

Sub GET_DETAILS(strPATH) 
 Dim arrHeaders(35) 
 Dim objShell 
 Dim objFolder 
 Dim i 

 Dim strFileName 
 Set objShell = CreateObject("Shell.Application") 
 Set objFolder = objShell.Namespace(strPATH) 

 For i = 0 to 34 
  arrHeaders(i) = objFolder.GetDetailsOf(objFolder.Items, i) 
 Next

 For Each strFileName in objFolder.Items 
  For i = 0 to 34 
   If objFolder.GetDetailsOf(strFileName, i) <> "" Then
    WScript.StdOut.WriteLine i & vbtab & arrHeaders(i) & ": " & objFolder.GetDetailsOf(strFileName, i) 
   End If
  Next

  WScript.StdOut.WriteLine "--------------------"
 Next
End Sub
