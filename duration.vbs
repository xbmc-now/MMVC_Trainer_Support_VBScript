Option Explicit
On Error Resume Next
'#������������������������������������������������������������������������������
'#��  wav�̒����`�F�b�N v1.0.1 (2022/12/08)
'#��  0.401�b�ȏ�܂��́A15.99�b�ȉ�
'#������������������������������������������������������������������������������
'#==============================================================================
'# �ݒ�l�������ɋL�ڂ���B
Dim Conf
Set Conf = CreateObject("Scripting.Dictionary")
'#-- [ ��{�ݒ� ] --------------------------------------------------------------


' �����t�H���_
Call Conf.Add("SearchFolder", "wav")

' �������ʃe�L�X�g
Call Conf.Add("SearchResult", "duration.txt")

'#------------------------------------------------------------------------------
Main()
If Err.Number <> 0 Then
	WScript.Echo "�G���[���������܂����B"
	WScript.Echo "�G���[�ԍ��F" & Err.Number & " " & "�G���[���e�F" & Err.Description
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

	Dim Folder
	' �J�����g�f�B���N�g���̌��s�t�H���_�Q��
	Set Folder = fso.GetFolder(ws.CurrentDirectory & "\" & Conf("SearchFolder"))

	Dim wrt
	Set wrt = fso.OpenTextFile(fso.BuildPath(ws.CurrentDirectory, Conf("SearchResult")), 2, True)

	SearchFolders ws, fso, wrt, Folder

	wrt.Close
	Set fso = Nothing
	Set ws  = Nothing
	Main = True
'#------------------------------------------------------------------------------
End Function

Sub SearchFolders(ws, fso, wrt, Folder)
	Dim File, subFolder, p, c, e, r, s
	' Folder���̃t�@�C����񋓂���
	For Each File in Folder.Files 

		If LCase(fso.GetExtensionName(File.name))="wav" Then ' �g���qwav�𔻒�
			p = Folder & "\" & File.name
			c = ws.CurrentDirectory & "\" & "ffprobe.exe -loglevel quiet -show_streams " & p
			Set e = ws.Exec("cmd /c " & c & " 2>&1 | findstr -i duration=")
			r = Split(Replace(e.StdOut.ReadAll, vbCrLf, ""), "=")
			s = CDbl(r(1))

			' 0.401�b�ȏ�܂��́A15.99�b�ȉ�
			If s < 0.401 Or s > 15.99 Then
				wrt.WriteLine "��" & File.name & ":" & s
				WScript.Echo "��" & File.name & ":" & s
			Else
				wrt.WriteLine File.name & ":" & s
				WScript.Echo File.name & ":" & s
			End If
		End If
	Next

	' Foler���̃t�H���_��񋓂���
	For Each subFolder in Folder.SubFolders 
		SearchFolders ws, fso, wrt, subFolder' �ċA�Ăяo��
	Next

End Sub
