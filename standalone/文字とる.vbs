'*******************************************************
'
' 文字トル Rev. 2005 Nov (Drag And Drop 対応版)
'
' 
'*******************************************************
Const cAppName = "文字トル Rev. 2005 Nov (Drag And Drop 対応版)"

Main()

Sub Main()
 Dim str

 n = WScript.Arguments.count
 If n = 0 Then
  str = GetSelectedFolderPath
  If str <> "" Then
   GenFileList str
  Else
   msgBox "ファイルリスト作成対象のフォルダを指定してください" & VbCrLf & VbCrLf & _
          "フォルダ内のファイルリストをExcelで作成して" & VbCrlf & _
          "指定した保存先に保存します" & Vbcrlf & _
          "ドラッグ＆ドロップでフォルダ指定してもOK",vbInformation,cAppName
  End If
  WScript.Quit
 end if

 For i = 0 To n - 1
  str = WScript.Arguments(i)
  If GetExtension(str)="" Then
   'ファイルリスト作成
   GenFileList str
  End If
 Next
End Sub

Function GetExtension(strFPath)
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
GetExtension = fso.GetExtensionName(strFPath)
End Function

Function SaveConvText(strConv,strSaveFPath)
Set fso = CreateObject("Scripting.FileSystemObject")
Dim ts
Const ForWriting =2

Set ts = fso.CreateTextFile(strSaveFPath, ForWriting, true)
ts.Write(strConv)
ts.Close
End Function

Function GetSelectedFolderPath()
  Dim ff
  Set ff = CreateObject("shell.application").browseforfolder(0, "ファイルリスト作成対象のフォルダを指定してください", 1)
  If Not ff Is Nothing Then
	If ff = "デスクトップ" Then
		GetSelectedFolderPath = "C:\Documents and Settings\All Users\デスクトップ"
		Exit Function	  
	End If	
	If ff.items.item Is Nothing Then
		Exit Function	  
	End If	
	GetSelectedFolderPath = ff.items.Item.Path
  End If
End Function



Sub GenerateExcelFileList(sFolder, SFileName, Ext)
	Dim TRGFolder
	Dim FSO

  Dim appExcel
  Dim ExcelBk
  Dim ExcelSheet

 Set FSO = CreateObject("Scripting.FileSystemObject")
 Set appExcel = CreateObject("Excel.Application")

  Set TRGFolder = FSO.GetFolder(sFolder)
  Set ExcelSheet = CreateObject("Excel.Sheet")
	
	Dim S
	Dim SubFolders
	Dim SubFolder
	Dim Files
	Dim File
	Dim arr()
	Dim Subarr()
	Dim i
	Dim Row
	Dim Col
	Dim SubdirFiles
	Dim SubdirFile
	Dim j
	Dim Ret

	Row = 1
	Col = 1
	ExcelSheet.ActiveSheet.Cells(Row, Col).Value = "リスト作成対象フォルダ :  "
	Row = Row + 1
	ExcelSheet.ActiveSheet.Cells(Row, Col).Value = TRGFolder.Path


	Set Files = TRGFolder.Files
	ReDim arr(Files.Count)


	Col = Col + 1
	ExcelSheet.ActiveSheet.Cells(Row, Col).Value = "フォルダ内のファイル数は  " & Files.Count

	If Files.Count <> 0 Then
		Row = Row + 1
		
		i = 0
		For Each File In Files
			If Ext=True Then
'				arr(i) = File.Name
				arr(i) = File.Attributes
			Else
				arr(i) = FSO.GetBaseName( File.Path)
			End If

			ExcelSheet.ActiveSheet.Cells(Row,Col - 1).Value = i + 1
			ExcelSheet.ActiveSheet.Cells(Row,Col).Value = arr(i)
			Row = Row + 1

			i=i+1
		Next

	End If

	Set SubFolders = TRGFolder.SubFolders

	If SubFolders.Count <> 0 Then
		'サブフォルダをリストに追加する
		InsertSubFolderList SubFolders, ExcelSheet, Row , Col, Ext
	End If


	'列幅を調整
	ExcelSheet.ActiveSheet.Rows("1:1").EntireColumn.AutoFit

	If SFileName <> False Then
		ExcelSheet.SaveAs SFileName
	End If

	msgbox "ファイルリストを作成しました" ,VbInformation + VbOKOnly ,cAppName
	appExcel.Quit
End Sub

Sub InsertSubFolderList(SubFolders, ExcelSheet, Row , Col, Ext)
	Dim SubFolder,SubSubFolders
	Dim Subarr()
	Dim SubdirFiles
	Dim SubdirFile
	Dim j
	Dim FSO
	Set FSO = CreateObject("Scripting.FileSystemObject")

	For Each SubFolder In SubFolders
		ExcelSheet.ActiveSheet.Cells(Row , Col).Value = SubFolder.Name
		
		Set SubdirFiles = SubFolder.Files
		ReDim Subarr(SubdirFiles.Count)
	
		ExcelSheet.ActiveSheet.Cells(Row,Col + 1).Value ="""" &  SubFolder.Name & """" & " フォルダ内のファイル数は  " & SubdirFiles.Count
		
		Row = Row + 1
		If SubdirFiles.Count <> 0 Then
	
			j = 0
			For Each SubdirFile In SubdirFiles
				If Ext=True Then
					Subarr(j) = SubdirFile.Name
				Else
					Subarr(j) = FSO.GetBaseName( SubdirFile.Path)
				End If

    				ExcelSheet.ActiveSheet.Rows(Row & ":" & Row).Insert
				ExcelSheet.ActiveSheet.Cells(Row,Col).Value = j + 1
				ExcelSheet.ActiveSheet.Cells(Row,Col + 1).Value = Subarr(j)
				Row = Row + 1
				j=j+1
			Next
		End If

	Set SubSubFolders = SubFolder.SubFolders

	If SubSubFolders.Count <> 0 Then
		Col = Col + 1
		InsertSubFolderList SubSubFolders, ExcelSheet, Row , Col, Ext
	End If

	Next
	Col = Col - 1
End Sub

Sub GenFileList(sFolder)
	Dim FileListPath

	Set objXL = CreateObject("Excel.Application")
	FileListPath = objXL.GetSaveAsFilename("FileList.xls","Excelファイル (*.xls),*.xls", 1, "ファイルリストの保存先を指定してください")

	If FileListPath = False Then
	 Msgbox "ファイルリストの作成がキャンセルされました",VbInformation +VbOkOnly ,cAppName
	 Exit Sub
	End If

	GenerateExcelFileList sFolder, FileListPath , True
End Sub
