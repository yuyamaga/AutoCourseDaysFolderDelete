Option Explicit

'エラー処理の確認
'ファイルを開いているときに削除が正常に行われるか確認

'-----------------------------------
' 定数宣言
'-----------------------------------
' TODO 実行環境用に変更する必要あり
' 削除対象フォルダ
Public Const conFolderPath = "C:\ProgramData\TOSHIBA\TEST"

' 差分が～日以上のフォルダを削除
Public Const conCourseDays = 0

' TODO 実行環境用に変更する必要あり
' ログファイル出力パス
Public Const conLogOutputFolderPath = "C:\ProgramData\TOSHIBA\TEST"

' TODO 実行環境用に変更する必要あり
' ログファイル出力パス
Public Const conLogOutputFilePath = "C:\ProgramData\TOSHIBA\TEST\log.txt"

' TODO 実行環境用に変更する必要あり
Public Const conMaxLogSize = 1000000

'-----------------------------------
' オブジェクト定義
'-----------------------------------
' ファイルシステムオブジェクト用
Public objFileSys

' サブフォルダ取得用
Dim objFolder

' 現在時刻保持用
Public strNowDateTime

' 現在時刻保持用
Public strNowDate

' サブフォルダ名称(ディレクトリパスあり)
Dim strFolderNamePath

' サブフォルダ名称(ディレクトリパスなし)
Dim strFolderName

' 日付形式(YYYY/MM/DD)に変更したフォルダ名保持用
Dim strFileDate

'-----------------------------------
' 各種オブジェクトの初期化
'-----------------------------------
' ファイルオブジェクト宣言
Set objFileSys = CreateObject("Scripting.FileSystemObject")

' ファイル情報取得用
Set objFolder = objFileSys.GetFolder(conFolderPath)

' 現在日数を取得
strNowDateTime = Now()

strNowDate = Year(Now()) & "/"

strNowDate = strNowDate & Right("0" & Month(Now()) , 2) & "/"

strNowDate = strNowDate & Right("0" & Day(Now()) , 2)

'-----------------------------------
' ディレクトリが存在するかで
' 以降の処理を行うか判定する
'-----------------------------------
If objFileSys.FolderExists( conFolderPath ) = False Then
	WriteLogFile( "ディレクトリパスに抽出対象フォルダが見つからなかったため、処理を終了" )
	WScript.Quit
End If

'-----------------------------------
' ディレクトリ内のフォルダ数を取得し、
' 以降の処理を行うか判定する
'-----------------------------------
If objFolder.SubFolders.Count < 1 Then
	WriteLogFile( "抽出対象フォルダ内にフォルダが一つも見つからなかったため、処理を終了" ) 
	WScript.Quit
End If

'-----------------------------------
' 直下の各フォルダ確認処理
'-----------------------------------
On Error Resume Next
For Each strFolderNamePath In objFolder.Subfolders
	'パスが入っていないフォルダ名称を取得
	strFolderName = strFolderNamePath.Name

	'フォルダ名称が指定形式(YYYYMMDD)のフォルダか判定する
	If CountLen(strFolderName) = 8 Then
		strFileDate = Mid(strFolderName, 1, 4) & "/"
		strFileDate = strFileDate & Mid(strFolderName, 5, 2) & "/"
		strFileDate = strFileDate & Mid(strFolderName, 7, 2)
		
		if DateDiff("d", strFileDate, strNowDate) >= conCourseDays then
			'フォルダ強制削除
			objFileSys.DeleteFolder strFolderNamePath,True
			
			WriteLogFile(strFolderName & "フォルダは" & DateDiff("d", strFileDate, strNowDate) & "日経過しているため、削除しました。")
		Else
			WriteLogFile(strFolderName & "フォルダは" & DateDiff("d", strFileDate, strNowDate) & "日経過、削除しません")
			'Wscript.Echo DateDiff("d", strFileDate, strNowDate) & "日経過、削除しません"
		End If
	Else
		'WScript.Echo "フォルダ名が指定形式(YYYYMMDD)でない為、次のフォルダ名を取得する"
	End If
Next
On Error Goto 0

'-----------------------------------
' フォルダ名称の桁数チェック用関数
' 8桁の数値かチェックする
'-----------------------------------
Function CountLen(ByVal data)
	Dim i
	Dim CheckData
	dim counter
	counter = 0
	for i = 1 To Len(data)
		'ASCIIコードに1文字ずつ変換してチェック
		CheckData = Asc(Mid(data, i, 1))
		If CheckData >= &H00 and CheckData <= &H7E then
			counter = counter + 1
		Else
			counter = counter + 2
		End If

	Next

	' 桁数を返却
	CountLen = counter
End Function

'-----------------------------------
' ログ出力用メソッド
'-----------------------------------
Private Sub WriteLogFile(strOutput)
	On Error Resume Next
    If objFileSys.FileExists(conLogOutputFilePath) Then
    	If objFileSys.GetFile(conLogOutputFilePath).Size >= conMaxLogSize Then
    		'フォルダ強制削除
			objFileSys.DeleteFolder conLogOutputFolderPath,True
		End If
		
    	' 追記保存
        With objFileSys.GetFile(conLogOutputFilePath).OpenAsTextStream(8)
            .WriteLine strNowDateTime & "_" & strOutput 
        End With
    Else
    	'If objFileSys.GetFile(conLogOutputFilePath).Size >= 100 Then
    	'End If
    	'ファイルが存在しない為、新規作成
        With objFileSys.CreateTextFile(conLogOutputFilePath, true)
            .WriteLine strNowDateTime & "_" & strOutput
        End With
    End If
    On Error Goto 0
End Sub


'-----------------------------------
' オブジェクト開放
'-----------------------------------
set objFileSys = Nothing
set objFolder = Nothing