Option Explicit
On Error Resume Next
'#������������������������������������������������������������������������������
'#��  fav�����R�s�[ v1.0.1 (2022/12/26)
'#��  fav�t�H���_�ɂ���e�L�X�g�t�@�C���Ɠ����̉����t�@�C����src�t�H���_�����out�t�H���_�ɃR�s�[
'#������������������������������������������������������������������������������
'#==============================================================================
'# �ݒ�l�������ɋL�ڂ���B
Dim Conf
Set Conf = CreateObject("Scripting.Dictionary")
'#-- [ ��{�ݒ� ] --------------------------------------------------------------


' �����t�H���_
Call Conf.Add("SearchFolder", "fav")

' ���t�@�C���t�H���_
Call Conf.Add("SourceFolder", "src")

' �o�̓t�H���_
Call Conf.Add("OutputFolder", "out")


'#------------------------------------------------------------------------------
Main()
If Err.Number <> 0 Then
	WScript.Echo "�G���[���������܂����B"
	WScript.Echo "�G���[�ԍ��F" & Err.Number & " �G���[���e�F" & Err.Description
End If

WScript.Echo vbLf & "�I�����܂����B�����L�[�������ƏI�����܂�"
WScript.StdIn.ReadLine
WScript.Quit

Function Main()
	Dim ws
	Set ws = CreateObject("WScript.Shell")

	'-----------------------------------------------------------------------------
	' CSCRIPT.EXE �Ŗ����ꍇ�ɐ؂�ւ���
	if Ucase(Right(WScript.FullName, 11)) <> "CSCRIPT.EXE" then
		' ������S�ēǂ݂����āA�X�y�[�X���܂܂�Ă����ꍇ�͍ēx�����n���ׂ� " �ň͂ݒ����Ă��܂�
		Dim strParam, a
		strParam = " "
		For a = 0 to Wscript.Arguments.Count - 1
			If instr(Wscript.Arguments(a), " ") < 1 Then
				strParam = strParam & Wscript.Arguments(a) & " "
			Else
				strParam = strParam & """" & Wscript.Arguments(a) & """ "
			End If
		Next

		' CSCRIPT.EXE �Ɉ����p���ŏI��
		Call ws.Run("cscript.exe //nologo """ & WScript.ScriptFullName & """" & strParam, 1)
		Set ws  = Nothing
		WScript.Quit

	end if
	'---------------------------------------------------------------------------
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")

	Dim re
	Set re = New RegExp
	re.Global = True

	Dim Folder
	' �J�����g�f�B���N�g���̌��s�t�H���_�Q��
	Set Folder = fso.GetFolder(ws.CurrentDirectory & "\" & Conf("SearchFolder"))


	SearchFolders ws, fso, re, Folder

	Set fso = Nothing
	Set ws  = Nothing
	Main = True
'#------------------------------------------------------------------------------
End Function

Sub SearchFolders(ws, fso, re, Folder)
	Dim File, ext, subFolder, p, o, c, e, r, CodecName, SampleRate, Channels, scStr
	' Folder���̃t�@�C����񋓂���
	For Each File in Folder.Files 
		ext = LCase(fso.GetExtensionName(File.name))

		If ext="txt" Then ' �g���qtxt�𔻒�

			p = ws.CurrentDirectory & "\" & Conf("SourceFolder") & "\" & fso.GetBaseName(File.name)
			o = ws.CurrentDirectory & "\" & Conf("OutputFolder") & "\" & fso.GetBaseName(File.name)

			If fso.FileExists(p & ".wav") Or fso.FileExists(p & ".ogg") Or fso.FileExists(p & ".mp3") Then
				If fso.FileExists(p & ".wav") Then
					fso.CopyFile p & ".wav", o & ".wav", True
					WScript.Echo File.name & "�R�s�[(wav)"
				End If
				If fso.FileExists(p & ".ogg") Then
					fso.CopyFile p & ".ogg", o & ".ogg", True
					WScript.Echo File.name & "�R�s�[(ogg)"
				End If
				If fso.FileExists(p & ".mp3") Then
					fso.CopyFile p & ".mp3", o & ".mp3", True
					WScript.Echo File.name & "�R�s�[(mp3)"
				End If
			Else
					WScript.Echo File.name & "�Ȃ�"
			End If


		End If
	Next

	' Foler���̃t�H���_��񋓂���
	For Each subFolder in Folder.SubFolders 
		SearchFolders ws, fso, re, subFolder' �ċA�Ăяo��
	Next

End Sub

