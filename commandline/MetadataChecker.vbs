'*******************************************************
'
' MetadataChecker Rev. D
'
' Date: modified 2011.12.16
' Author: takahashi hideaki
'*******************************************************

' 要追加
' """の検索、置換
' $"および"$の検索置換

' 使用上の注意
'
' メタデータの列数（フィールド数）を変更する場合は、
' 下記の定数を修正する
'
'*******************************************************
' メタデータの列数（フィールド数）設定用の定数
'*******************************************************
Function GetCols(sMetadataType)
Dim sUcase_Type
sUcase_Type = UCase(sMetadataType)
Select Case sUcase_Type
Case "TIR_SUPPLIES"
 GetCols = 12
Case "IPC_SPARES"
 GetCols = 10
Case "TIR_TOOLS","Support eq/Tools"
 GetCols = 13
Case "TIR_ORGANIZATIONS","Supplies"
 GetCols = 5
Case "TIR_ENTERPRISE","TIR_Enterprise"
 GetCols = 18
Case "TIR_CIRCUIT_BREAKERS","Circuit Breakers"
 GetCols = 8
Case "TIR_ZONES","Zones"
 GetCols = 8
Case "TIR_ACCESS_POINTS","Access Points","TIR_ACCESSPOINTS","AccessPoints"
 GetCols = 10
Case "AUTHOR"
 GetCols = 17
'Case "INTEGRATION"
' GetCols = 12
Case "ILLUSTRATION"
 GetCols = 9
Case Else
 GetCols = 12
End Select

End Function


'*******************************************************
' メイン呼び出し (消さないで！)
'*******************************************************
Main()


'*******************************************************
'　定数
'*******************************************************
Const cAppName = "Saab送付用CSVチェッカー MetadataChecker Rev. E"
Const cHowto = "Saab送付用CSVファイルをドラッグ＆ドロップしてください"
Const cHowtoDetail = ""
Const cTSVOnly = "拡張子が.csvのファイルのみドラッグ＆ドロップしてください"

' metadataの2列目にファイル名
Const cFileNameColNum = 1
' metadataの1列目はCategory
Const cFileCategoryColNum = 0
' metadataの6列目はIntegration DB Type
Const cDBTypeColNum = 5

'*******************************************************
' メイン
'*******************************************************
Sub Main()
 Dim str
 Dim sFname
 Dim MetadataType 
 Dim ArrFPath()
 Dim ArrFName()
 Dim ExistMetadata 
 Dim IncludingNotTSVFormat 

 ' ドラッグ＆ドロップされたファイルの個数を取得(= n)
 n = WScript.Arguments.count

 ' ドラッグ＆ドロップされたファイルがない
 If n = 0 Then
   ' 使用方法を表示して終了（何もしない）
   msgBox "Saab送付用CSVファイルのフォーマットをチェックします" & VbCrLf & _
	cHowto & VbCrLf & cKoujityu ,vbInformation,cAppName
  WScript.Quit
 End If

 '初期化
 ReDim ArrFPath(n)
 ReDim ArrFName(n)
 ExistMetadata = False
 IncludingNotTSVFormat = False

 ' ドラッグ＆ドロップされたファイルのフルパスとファイル名を格納
 For i = 0 To n - 1
  'ドラッグされたファイルのフルパスを取得
  ArrFPath(i) = WScript.Arguments(i)
  'ファイル名を取得
  ArrFName(i) = GetFileName(ArrFPath(i))

  ' メタデータの有無を判定
  If IsMetadataFile(ArrFName(i)) = True Then
   ExistMetadata = True
  End If

   ' csv以外の拡張子があるかを判定
  If GetExtension(ArrFPath(i))<>"csv" Then
   IncludingNotTSVFormat = True
  End If

 Next


 ' n=1：1ファイルのみドラッグ＆ドロップされた場合
 If n = 1 Then

   ' 拡張子がcsv以外である場合はチェック対象外
   If IncludingNotTSVFormat = True Then
    msgBox cHowto,vbInformation,cAppName
    Exit Sub
   End If

   ' メタデータかどうか判定
   ' ファイル名に大文字小文字を区別せずに"metadata"の文字列があれば、メタデータとみなす
   ' メタデータでない場合
   If ExistMetadata = False Then
    ' --> メタデータと送付データをドラッグ＆ドロップしてもらう
       'integrationの場合はフォーマットチェックする
        sFName =ArrFName(0)
      if  GetMetadataType(sFName)="INTEGRATION" then
        MetadataType = GetIntegrationFileType (ArrFName(0))
        TSVFileCheck ArrFPath(0), GetCols(MetadataType),MetadataType
      else
       msgBox cHowtoDetail,vbInformation,cAppName
       ' -->処理せずに終了
       Exit Sub
      end if
   ' メタデータのみドラッグ＆ドロップされた場合
   Else
    ' --> メタデータチェックする
    MetadataCheck ArrFPath(0), ArrFName(0)
    Exit Sub
   End If

 ' n>2：2ファイル以上ドラッグ＆ドロップされた場合
 Else

   ' メタデータが無ければ、送付ファイルもチェックしない
   If ExistMetadata = False Then
    msgBox cHowtoDetail,vbInformation,cAppName
    Exit Sub

   ' 拡張子がcsv以外である場合はチェック対象外
   ElseIf IncludingNotTSVFormat = True Then
    msgBox cTSVOnly ,vbInformation,cAppName
    Exit Sub

   Else
   ' メタデータがあれば、メタデータと送付ファイルのチェックをする
    For i = 0 To n - 1

     ' メタデータの場合
     If IsMetadataFile(ArrFName(i)) = True Then
      'メタデータチェック
      MetadataCheck ArrFPath(i), ArrFName(i)
'     'メタデータ以外のTIR等のTSVフォーマット
     Else
'      'TIR等のTSVフォーマットが適切かどうかをチェックする
   'メタデータのフォーマットをチェックする
   MetadataType = GetIntegrationFileType (ArrFName(i))

      TSVFileCheck ArrFPath(i), GetCols(MetadataType),MetadataType
     End If
    Next


    Exit Sub
   End If

 End If

End Sub


'*******************************************************
' メタデータチェックする
' MetadataCheck
'*******************************************************
Sub MetadataCheck(sFPath ,sFname)

   'メタデータのフォーマットをチェックする
   MetadataType = GetMetadataType(sFname)

   'メタデータのファイル命名規則に合っているかチェック
   If IsApplicableMetadataName(sFname)=False Then
    msgbox "このメタデータはファイル命名規則に合致していません！" ,vbExclamation,cAppName
   Else
    msgbox "このメタデータはファイル命名規則に合致しています" ,vbInformation,cAppName
   End If

   'TSVのフォーマットチェックを実行
   MetadataTSVFormatCheck sFPath, GetCols(MetadataType)
End Sub



'*******************************************************
' TSVのフォーマットチェックを実行
' MetadataTSVFormatCheck
' 戻り値：
' True：不備なし
' False：不備あり
'*******************************************************
Function MetadataTSVFormatCheck(strFPath, fldNum)

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
MetadataTSVFormatCheck=True

Dim metadata_file
Dim sFolder
Dim i
i=0

Dim sLogFPath
sLogFPath = GetsErrLogFPath(strFPath)

Dim j
Set metadata_file =fso.OpenTextFile(strFPath)
sFolder = fso.GetParentFolderName(strFPath)
 
Dim CurrentLineText
Dim TSVFields
Dim sErrMsg
Dim sIntegrationFPath
Dim ExistEOF
ExistEOF=False
CurrentLineText=""
	Do Until (metadata_file.AtEndOfStream)
		i=i+1
		CurrentLineText = metadata_file.ReadLine
		'WScript.Echo CurrentLineText
			
		If Instr(1,CurrentLineText,"EOF",1) Then
			ExistEOF=True
			Exit Do
		End if

		' 1行の中の区切り文字(Tab)数が規定に合わない場合メッセージを表示する
		If GetSeparateCharCnt(fldNum)<>CntStr(CurrentLineText ,"$") Then

			sErrMsg = sErrMsg & vbCrLf & "Line " & i & " ：メタデータの区切り文字に不備があります！" & vbcrlf & _
			CurrentLineText & vbcrlf & "---> セル内の改行がないか、あるいは、指定されたフィールド数になっているかをExcel上で確認してください。" & vbcrlf & vbcrlf
'			msgbox "メタデータの区切り文字に不備があります！" & " メタデータの" & i & "行目" & vbcrlf & vbcrlf & _
'			CurrentLineText & vbcrlf & vbcrlf & "セル内の改行がないか、指定されたフィールド数になっているかをExcel上で確認してください。" ,vbCritical,cAppName

			MetadataTSVFormatCheck=False
			Exit Do
		End if

		' Tab区切りで分ける
		TSVFields = Split(CurrentLineText ,"$")
		j=0

		If GetMetadataType(TSVFields(cFileCategoryColNum)) ="INTEGRATION" Then
		 msgbox TSVFields(cFileNameColNum) & " のフォーマットをチェックします!" & VbCrLf & "File Type : " & "INTEGRATION",vbInformation,cAppName
		 ' INTEGRATION fileのフォーマットをチェック
		 sIntegrationFPath = sFolder & "\" & TSVFields(cFileNameColNum)
		 TSVFileCheck sIntegrationFPath,GetCols(TSVFields(cDBTypeColNum)),TSVFields(cDBTypeColNum)
		End If
		For j=0 to CntStr(CurrentLineText ,"$") -1
		' ASCII以外の文字コードが含まれる場合にメッセージを表示する
		If Ascii_chk(TSVFields(j)) = True Then
			If Len(TSVFields(j)) <> LenByte(TSVFields(j)) Then
				sErrMsg = sErrMsg & "Line " & i & " ：全角文字が含まれています ---> " & TSVFields(j) & vbcrlf
'				msgbox "全角文字が含まれています" & " メタデータの" & i & "行目" & vbcrlf & vbcrlf & _
'				TSVFields(j) ,vbExclamation,cAppName
			Else
				sErrMsg = sErrMsg & "Line " & i & " ：ASCIIコード以外の文字が含まれています ---> " & TSVFields(j) & vbcrlf
'				msgbox "ASCIIコード以外の文字が含まれています" & " メタデータの" & i & "行目" & vbcrlf & vbcrlf & _
'				TSVFields(j) ,vbExclamation,cAppName
			End If

			MetadataTSVFormatCheck=False
		End If
		Next

	Loop


If MetadataTSVFormatCheck=True Then
   If ExistEOF=False then
      sErrMsg = sErrMsg & VbCrLf & "最終行にEOFが見つかりませんでした"
      msgbox "最終行にEOFが見つかりませんでした",vbExclamation,cAppName
   End If
	sErrMsg = sErrMsg & "メタデータのチェックを終了しました。" & vbCrLf & "メタデータに使用できない文字は含まれていませんでした。" & VbCrLf & strFPath
	SaveLogText sErrMsg, sLogFPath
	Msgbox sErrMsg,vbInformation,cAppName
Else
	sErrMsg = sErrMsg & "メタデータのチェックを終了しました。" & vbCrLf & "メタデータに不備があります。" & VbCrLf & strFPath
	SaveLogText sErrMsg, sLogFPath
	Msgbox "メタデータのチェックを終了しました。"  & vbcrlf & "メタデータに不備があります" & vbcrlf & "ログファイルを確認して下さい" & VbCrLf & sLogFPath,vbCritical,cAppName
End If

metadata_file.Close
End Function


'*******************************************************
' TSVのフォーマットチェックを実行
' TSVFileCheck
' 戻り値：
' True：不備なし
' False：不備あり
'*******************************************************
Function TSVFileCheck(strFPath, fldNum, sFType)

Dim fso
Dim ret
Set fso = CreateObject("Scripting.FileSystemObject")
TSVFileCheck=True

ret = fso.FileExists( strFPath )

If ret = False Then
 msgbox sFType & " :" &  strFPath & VbCrLf & "がメタデータと同一フォルダに存在しません！" & VbCrLf & "メタデータと同一フォルダに" & sFType & "を保存して実行してください",vbInformation,cAppName
 Exit Function
End If

Dim tsv_file
Dim i
i=0

Dim sLogFPath
sLogFPath = GetsErrLogFPath(strFPath)

Dim j
Set tsv_file =fso.OpenTextFile(strFPath)

Dim CurrentLineText
Dim TSVFields
Dim sErrMsg
CurrentLineText=""
Dim ExistEOF
ExistEOF=False

	Do Until (tsv_file.AtEndOfStream)
		i=i+1
		CurrentLineText = tsv_file.ReadLine
		'WScript.Echo CurrentLineText

		If Instr(1,CurrentLineText,"EOF",1) Then
			ExistEOF=True
			Exit Do
		End if
			
'tmp=CntStr(CurrentLineText ,"$")
'tmp2=GetSeparateCharCnt(fldNum)
'msgbox tmp
'msgbox tmp2
	
		' 1行の中の区切り文字(Tab)数が規定に合わない場合メッセージを表示する
		If GetSeparateCharCnt(fldNum)<>CntStr(CurrentLineText ,"$") Then
			sErrMsg = sErrMsg & vbCrLf & "Line " & i & " ：区切り文字に不備があります！" & vbcrlf & _
			CurrentLineText & vbcrlf & "---> 行末の$が付与されていないか、セル内の改行がないか、指定されたフィールド数になっているかをExcel上で確認してください。" & vbcrlf & vbcrlf
'			msgbox "区切り文字に不備があります！" & " " & i & "行目" & vbcrlf & vbcrlf & _
'			CurrentLineText & vbcrlf & vbcrlf & " 行末の$が付与されていないか、セル内の改行がないか、指定されたフィールド数になっているかをExcel上で確認してください。" ,vbCritical,cAppName

			TSVFileCheck=False
			Exit Do
		End if

		' Tab区切りで分ける
		TSVFields = Split(CurrentLineText ,"$")
		j=0
		
		For j=0 to CntStr(CurrentLineText ,"$") -1
		' ASCII以外の文字コードが含まれる場合にメッセージを表示する
		If Ascii_chk(TSVFields(j)) = True Then
			If Len(TSVFields(j)) <> LenByte(TSVFields(j)) Then
				sErrMsg = sErrMsg & "Line " & i & " ：全角文字が含まれています ---> " & TSVFields(j) & vbcrlf
'				msgbox "全角文字が含まれています" & " " & i & "行目" & vbcrlf & vbcrlf & _
'				TSVFields(j) ,vbExclamation,cAppName
			Else
				sErrMsg = sErrMsg & "Line " & i & " ：ASCIIコード以外の文字が含まれています ---> " & TSVFields(j) & vbcrlf
'				msgbox "ASCIIコード以外の文字が含まれています" & " " & i & "行目" & vbcrlf & vbcrlf & _
'				TSVFields(j) ,vbExclamation,cAppName
			End If

			TSVFileCheck=False
		End If
		Next

	Loop

If ExistEOF=False then
    sErrMsg = sErrMsg & VbCrLf & "最終行にEOFが見つかりませんでした"
msgbox "最終行にEOFが見つかりませんでした",vbExclamation,cAppName
End If

If TSVFileCheck=True Then
	sErrMsg = strFPath & VbCrLf & "のチェックを終了しました。" & vbCrLf & sFType & " Fileに使用できない文字は含まれていませんでした。" 
	SaveLogText sErrMsg, sLogFPath
	Msgbox sErrMsg ,vbInformation,cAppName
Else
	SaveLogText sErrMsg, sLogFPath
	Msgbox strFPath & VbCrLf & "のチェックを終了しました。"  & vbcrlf & sFType & " Fileに不備があります" & vbcrlf & "ログファイルを確認して下さい" & VbCrLf & sLogFPath,vbCritical,cAppName

End If

tsv_file.Close
End Function


'*******************************************************
' metadata用ファイルかどうかの判定
' 判定方法：ファイル名に"metadata"の文字列が含まれている
'*******************************************************
Function IsMetadataFile(s)
Dim sUcase_FName
IsMetadataFile = False
sUcase_FName = UCase(s)
If InStr(sUcase_FName,"METADATA") > 0 Then
IsMetadataFile = True
End If
End Function

'*******************************************************
' メタデータのファイル名規則に合っているかどうかを判定する
' IsApplicableMetadataName
' True：不備なし
' False：不備あり
'*******************************************************
Function IsApplicableMetadataName(sFName)
cTIRFiles = "Integration"

If GetFileCategory(sFName)="UNKNOWN" Then
 IsApplicableMetadataName =False
 Exit Function
End if
If IsYearDateFormat(sFName)=False Then
 IsApplicableMetadataName =False
 Exit Function
End if


IsApplicableMetadataName = True
End Function

'*******************************************************
' メタデータのFile Categoryを取得する
'*******************************************************
Function GetFileCategory(sFName)
Dim sUcase_FName
GetFileCategory = "UNKNOWN"
sUcase_FName = UCase(sFName)
If InStr(1, sUcase_FName,"INTEGRATION") > 0 Then
GetFileCategory= "INTEGRATION"
End If
If InStr(1, sUcase_FName,"AUTHOR") Then
GetFileCategory = "AUTHOR"
End If
End Function




'*******************************************************
' ASCIIコード以外の文字コードが含まれているかマッチング
' 戻り値：
' True：ASCIIコード以外の文字コードが含まれている
' False：含まれない
'*******************************************************
Function Ascii_chk(s)
Dim objRE
Set objRE = new RegExp
objRE.IgnoreCase = True
objRE.pattern = "[^\x20-\x7E]"
Ascii_chk = objRE.Test(s)

Set objRE = Nothing
End Function




'*******************************************************
' メタデータの種類を識別する
'*******************************************************
Function GetMetadataType(sFName)
Dim sUcase_FName
GetMetadataType = "Unknown"
sUcase_FName = UCase(sFName)

If InStr(1, sUcase_FName,"ILLUSTRATION") > 0 Then
  GetMetadataType= "ILLUSTRATION"
  Exit Function
End If
If InStr(1, sUcase_FName,"AUTHOR") Then
 GetMetadataType = "AUTHOR"
End If
If InStr(1, sUcase_FName,"INTEGRATION") > 0 Then
  GetMetadataType= "INTEGRATION"
End If
End Function

'*******************************************************
' IntegrationFileの種類を識別する
'*******************************************************
Function GetIntegrationFileType(sFName)
Dim sUcase_FName
GetIntegrationFileType = "Unknown"
sUcase_FName = UCase(sFName)

 If InStr(1, sUcase_FName,"IPC_SPARES") > 0 Then
  GetIntegrationFileType= "IPC_SPARES"
 End If
 If InStr(1, sUcase_FName,"TIR_SUPPLIES") > 0 Then
  GetIntegrationFileType= "TIR_SUPPLIES"
 End If
 If InStr(1, sUcase_FName,"TIR_TOOLS") > 0 Then
  GetIntegrationFileType= "TIR_TOOLS"
 End If
 If InStr(1, sUcase_FName,"TIR_ORGANIZATIONS") > 0 Then
  GetIntegrationFileType= "TIR_ORGANIZATIONS"
 End If
 If InStr(1, sUcase_FName,"TIR_ENTERPRISE") > 0 Then
  GetIntegrationFileType= "TIR_ENTERPRISE"
 End If
 If InStr(1, sUcase_FName,"TIR_CIRCUIT_BREAKERS") > 0 Then
  GetIntegrationFileType= "TIR_CIRCUIT_BREAKERS"
 End If
 If InStr(1, sUcase_FName,"TIR_ZONES") > 0 Then
  GetIntegrationFileType= "TIR_ZONES"
 End If
 If InStr(1, sUcase_FName,"TIR_ACCESS_POINTS") > 0 Then
  GetIntegrationFileType= "TIR_ACCESS_POINTS"
 End If
End Function


'*******************************************************
' yyyymmdd_hhmmの文字列形式になっているかマッチング
' 戻り値：
' True：yyyymmdd_hhmmの文字列形式になっている
' False：yyyymmdd_hhmmの文字列形式でない
'*******************************************************
Function IsYearDateFormat(s)
Dim objRE
Set objRE = new RegExp
objRE.IgnoreCase = True
' yyyymmdd_hhmmの文字列形式になっているかマッチング
objRE.pattern = "([0-9]{4})([0-9]{2})([0-9]{2})_([0-9]{2})([0-9]{2})"
IsYearDateFormat = objRE.Test(s)
Set objRE = Nothing
End Function

'*******************************************************
' メタデータの１行内のタブ想定数
'*******************************************************
Function GetSeparateCharCnt(TTLCols)
	GetSeparateCharCnt = TTLCols
'行末に$を付加することになったため変更
'	GetSeparateCharCnt = TTLCols-1
End Function

'*******************************************************
'アスキー文字バイト数取得
'*******************************************************
Function LenByte(ByVal s)
    Dim c, i, k    
c = 0    
For i = 0 To Len(s) - 1        
k = Mid(s, i + 1, 1)        
If (Asc(k) And &HFF00) = 0 Then
            c = c + 1        
Else
            c = c + 2        
End If
Next
    LenByte = c
End Function


'*******************************************************
' ログファイル書き出し
'*******************************************************
Sub SaveLogText(sErrLog,sErrLogFPath)
Set fso = CreateObject("Scripting.FileSystemObject")
Dim ts
Const ForWriting =2

Set ts = fso.CreateTextFile(sErrLogFPath, ForWriting, true)
ts.Write(sErrLog)
ts.Close
End Sub



'*******************************************************
' ログファイル書き出し先パスを取得
'*******************************************************
Function GetsErrLogFPath(sTSVFPath)
Dim myFname
Dim FS
Dim ret

'ファイルシステムオブジェクトを生成
Set FS = CreateObject("Scripting.FileSystemObject")

'保存するログファイルの拡張子は"log"にした
'myFname = Replace(sTSVFPath, "tsv", "log")
myFname = Replace(sTSVFPath, "csv", "log")
ret = FS.FileExists( myFname )

If ret = True Then
	GetsErrLogFPath = Replace(myFname , ".log", "_1.log")
Else
	GetsErrLogFPath = myFname
End If
End Function

'*******************************************************
' 拡張子取得
'*******************************************************
Function GetExtension(strFPath)
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
GetExtension = fso.GetExtensionName(strFPath)
End Function

'*******************************************************
' ファイル名取得
'*******************************************************
Function GetFileName(strFPath)
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
GetFileName = fso.GetFileName(strFPath)
End Function


'*******************************************************
' 文字列の出現数をカウント
'*******************************************************
Function CntStr(s, org)
  dim i
  dim j
  dim k


  k = len(org)
  i = 1
  j = 0
  do
    i = instr(i, s, org)
    if i > 0 then
     i = i + k
     j = j + 1
    end if
  loop until i = 0

  CntStr = j

End Function

