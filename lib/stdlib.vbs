'================================================================
' Common Liblary
'
'
'*******************************************************
Const cNoFolderOrFile ="対象フォルダが指定されていません！"

Const cUsage02 = "対象フォルダもしくはファイルをこのアイコン上に"
Const cUsage03 = "ドラッグ & ドロップして処理をおこなえます。"
'Const cUsage02 = "YOU CAN ALSO SPECIFY THE FOLDER BY DRAG AND DROP ONTO THIS FILE."

Const cAPP0001 = "TestSubProcedure"
Const cAPP0002 = "PrepareSaabBOM"
Const cAPP0003 = "GetPDFPageCount"




Function GetSelectedFolderPath()
  Dim ff
  Set ff = CreateObject("shell.application").browseforfolder(0, cIndicatePurpose & VbCrLf & cSelectFolderOrFile & VbCrLf & cUsage02 & VbCrLf & cUsage03, 1)
  If Not ff Is Nothing Then
	If ff = "Desktop" Then
		GetSelectedFolderPath = "C:\Documents and Settings\All Users\Desktop"
		Exit Function	  
	End If	
	If ff.items.item Is Nothing Then
		Exit Function	  
	End If	
	GetSelectedFolderPath = ff.items.Item.Path
  End If
End Function

Function GetExtension(strFPath)
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
GetExtension = fso.GetExtensionName(strFPath)
End Function

Function IsExcelFile(sExtension)
Select Case sExtension
Case "xls","xlsx","xlsm"
 IsExcelFile = True
Case Else
 IsExcelFile = False
End Select
End Function

Function IsPDFFile(sExtension)
Select Case sExtension
Case "pdf"
 IsPDFFile = True
Case Else
 IsPDFFile = False
End Select
End Function


Sub ActionWhenIconClicked()
 Dim strSelectedFolderPath
  strSelectedFolderPath = GetSelectedFolderPath
  If strSelectedFolderPath <> "" Then
   Call StartAppWithDragAndDroppedPath(strSelectedFolderPath, cAppMain)
  Else
   'No draged file or folder???
   msgBox cNoFolderOrFile,vbExclamation,cAppName
  End If
End Sub

Function StartAppWithDragAndDroppedPath(sDragAndDroppedPath, sAppMain)
 Dim sFunctionName0001
 Dim sFunctionName0002
 Dim sFunctionName0003

 sFunctionName0001=cAPP0001
 sFunctionName0002=cAPP0002
 sFunctionName0003=cAPP0003

 Select Case sAppMain
 Case sFunctionName0001
  TestSubProcedure(sDragAndDroppedPath)
  StartAppWithDragAndDroppedPath = False
 Case sFunctionName0002
  PrepareSaabBOM(sDragAndDroppedPath)
  StartAppWithDragAndDroppedPath = False
 Case sFunctionName0003
  StartAppWithDragAndDroppedPath = GetPDFPageCount(sDragAndDroppedPath)
 End Select
End Function
