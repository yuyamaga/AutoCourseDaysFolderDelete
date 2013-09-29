Option Explicit

'�G���[�����̊m�F
'�t�@�C�����J���Ă���Ƃ��ɍ폜������ɍs���邩�m�F

'-----------------------------------
' �萔�錾
'-----------------------------------
' TODO ���s���p�ɕύX����K�v����
' �폜�Ώۃt�H���_
Public Const conFolderPath = "C:\ProgramData\TOSHIBA\TEST"

' �������`���ȏ�̃t�H���_���폜
Public Const conCourseDays = 0

' TODO ���s���p�ɕύX����K�v����
' ���O�t�@�C���o�̓p�X
Public Const conLogOutputFolderPath = "C:\ProgramData\TOSHIBA\TEST"

' TODO ���s���p�ɕύX����K�v����
' ���O�t�@�C���o�̓p�X
Public Const conLogOutputFilePath = "C:\ProgramData\TOSHIBA\TEST\log.txt"

' TODO ���s���p�ɕύX����K�v����
Public Const conMaxLogSize = 1000000

'-----------------------------------
' �I�u�W�F�N�g��`
'-----------------------------------
' �t�@�C���V�X�e���I�u�W�F�N�g�p
Public objFileSys

' �T�u�t�H���_�擾�p
Dim objFolder

' ���ݎ����ێ��p
Public strNowDateTime

' ���ݎ����ێ��p
Public strNowDate

' �T�u�t�H���_����(�f�B���N�g���p�X����)
Dim strFolderNamePath

' �T�u�t�H���_����(�f�B���N�g���p�X�Ȃ�)
Dim strFolderName

' ���t�`��(YYYY/MM/DD)�ɕύX�����t�H���_���ێ��p
Dim strFileDate

'-----------------------------------
' �e��I�u�W�F�N�g�̏�����
'-----------------------------------
' �t�@�C���I�u�W�F�N�g�錾
Set objFileSys = CreateObject("Scripting.FileSystemObject")

' �t�@�C�����擾�p
Set objFolder = objFileSys.GetFolder(conFolderPath)

' ���ݓ������擾
strNowDateTime = Now()

strNowDate = Year(Now()) & "/"

strNowDate = strNowDate & Right("0" & Month(Now()) , 2) & "/"

strNowDate = strNowDate & Right("0" & Day(Now()) , 2)

'-----------------------------------
' �f�B���N�g�������݂��邩��
' �ȍ~�̏������s�������肷��
'-----------------------------------
If objFileSys.FolderExists( conFolderPath ) = False Then
	WriteLogFile( "�f�B���N�g���p�X�ɒ��o�Ώۃt�H���_��������Ȃ��������߁A�������I��" )
	WScript.Quit
End If

'-----------------------------------
' �f�B���N�g�����̃t�H���_�����擾���A
' �ȍ~�̏������s�������肷��
'-----------------------------------
If objFolder.SubFolders.Count < 1 Then
	WriteLogFile( "���o�Ώۃt�H���_���Ƀt�H���_�����������Ȃ��������߁A�������I��" ) 
	WScript.Quit
End If

'-----------------------------------
' �����̊e�t�H���_�m�F����
'-----------------------------------
On Error Resume Next
For Each strFolderNamePath In objFolder.Subfolders
	'�p�X�������Ă��Ȃ��t�H���_���̂��擾
	strFolderName = strFolderNamePath.Name

	'�t�H���_���̂��w��`��(YYYYMMDD)�̃t�H���_�����肷��
	If CountLen(strFolderName) = 8 Then
		strFileDate = Mid(strFolderName, 1, 4) & "/"
		strFileDate = strFileDate & Mid(strFolderName, 5, 2) & "/"
		strFileDate = strFileDate & Mid(strFolderName, 7, 2)
		
		if DateDiff("d", strFileDate, strNowDate) >= conCourseDays then
			'�t�H���_�����폜
			objFileSys.DeleteFolder strFolderNamePath,True
			
			WriteLogFile(strFolderName & "�t�H���_��" & DateDiff("d", strFileDate, strNowDate) & "���o�߂��Ă��邽�߁A�폜���܂����B")
		Else
			WriteLogFile(strFolderName & "�t�H���_��" & DateDiff("d", strFileDate, strNowDate) & "���o�߁A�폜���܂���")
			'Wscript.Echo DateDiff("d", strFileDate, strNowDate) & "���o�߁A�폜���܂���"
		End If
	Else
		'WScript.Echo "�t�H���_�����w��`��(YYYYMMDD)�łȂ��ׁA���̃t�H���_�����擾����"
	End If
Next
On Error Goto 0

'-----------------------------------
' �t�H���_���̂̌����`�F�b�N�p�֐�
' 8���̐��l���`�F�b�N����
'-----------------------------------
Function CountLen(ByVal data)
	Dim i
	Dim CheckData
	dim counter
	counter = 0
	for i = 1 To Len(data)
		'ASCII�R�[�h��1�������ϊ����ă`�F�b�N
		CheckData = Asc(Mid(data, i, 1))
		If CheckData >= &H00 and CheckData <= &H7E then
			counter = counter + 1
		Else
			counter = counter + 2
		End If

	Next

	' ������ԋp
	CountLen = counter
End Function

'-----------------------------------
' ���O�o�͗p���\�b�h
'-----------------------------------
Private Sub WriteLogFile(strOutput)
	On Error Resume Next
    If objFileSys.FileExists(conLogOutputFilePath) Then
    	If objFileSys.GetFile(conLogOutputFilePath).Size >= conMaxLogSize Then
    		'�t�H���_�����폜
			objFileSys.DeleteFolder conLogOutputFolderPath,True
		End If
		
    	' �ǋL�ۑ�
        With objFileSys.GetFile(conLogOutputFilePath).OpenAsTextStream(8)
            .WriteLine strNowDateTime & "_" & strOutput 
        End With
    Else
    	'If objFileSys.GetFile(conLogOutputFilePath).Size >= 100 Then
    	'End If
    	'�t�@�C�������݂��Ȃ��ׁA�V�K�쐬
        With objFileSys.CreateTextFile(conLogOutputFilePath, true)
            .WriteLine strNowDateTime & "_" & strOutput
        End With
    End If
    On Error Goto 0
End Sub


'-----------------------------------
' �I�u�W�F�N�g�J��
'-----------------------------------
set objFileSys = Nothing
set objFolder = Nothing