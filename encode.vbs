Option Explicit
On Error Resume Next
'#������������������������������������������������������������������������������
'#��  wav�G���R�[�h v1.0.1 (2022/12/08)
'#��  src�t�H���_�ɂ��鉹���t�@�C����wav�t�H���_�ɃG���R�[�h
'#������������������������������������������������������������������������������
'#==============================================================================
'# �ݒ�l�������ɋL�ڂ���B
Dim Conf
Set Conf = CreateObject("Scripting.Dictionary")
'#-- [ ��{�ݒ� ] --------------------------------------------------------------


' �����t�H���_
Call Conf.Add("SearchFolder", "src")

' �o�̓t�H���_
Call Conf.Add("OutputFolder", "wav")

' �������ʃe�L�X�g
Call Conf.Add("SearchResult", "encode.txt")

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

	Dim wrt
	Set wrt = fso.OpenTextFile(fso.BuildPath(ws.CurrentDirectory, Conf("SearchResult")), 2, True)

	SearchFolders ws, fso, re, wrt, Folder

	wrt.Close
	Set fso = Nothing
	Set ws  = Nothing
	Main = True
'#------------------------------------------------------------------------------
End Function

Sub SearchFolders(ws, fso, re, wrt, Folder)
	Dim File, ext, subFolder, p, o, c, e, r, CodecName, SampleRate, Channels, scStr
	' Folder���̃t�@�C����񋓂���
	For Each File in Folder.Files 
		ext = LCase(fso.GetExtensionName(File.name))

		If ext="wav" Or ext="ogg" Or ext="mp3" Then ' �g���qwav, ogg, mp3�𔻒�

			p = Folder & "\" & File.name
			c = ws.CurrentDirectory & "\" & "ffprobe.exe -loglevel quiet -show_streams " & p
			Set e = ws.Exec(c)
			r = e.StdOut.ReadAll

			CodecName  = getWavInfo(re, r, "codec_name")
			SampleRate = getWavInfo(re, r, "sample_rate")
			Channels   = getWavInfo(re, r, "channels")

			o = ws.CurrentDirectory & "\" & Conf("OutputFolder") & "\" & fso.GetBaseName(File.name) & ".wav"
			If ext = "wav" AND CodecName = "pcm_s16le" AND SampleRate = "24000" AND Channels = "1" Then
				' 24000Hz 16bit 1ch��wav�t�@�C���ł���΁A�G���R�[�h�����R�s�[
				fso.CopyFile p, o, True
				WScript.Echo File.name & "�R�s�["
				wrt.WriteLine File.name & "�R�s�["
			Else
				c = ws.CurrentDirectory & "\" & "ffmpeg.exe -y -i " & p & " -acodec pcm_s16le -ar 24000 -ac 1 " & o
				Set e = ws.Exec(c)
				'r = e.StdErr.ReadAll
				WScript.Echo File.name & "�G���R�[�h"
				wrt.WriteLine File.name & "�G���R�[�h"
			End If
		End If
	Next

	' Foler���̃t�H���_��񋓂���
	For Each subFolder in Folder.SubFolders 
		SearchFolders ws, fso, re, wrt, subFolder' �ċA�Ăяo��
	Next

End Sub

Function getWavInfo(re, str, item)
	Dim m, s
	re.Pattern = item & "=.+\n"
	Set m = re.Execute(str)
	s = Split(Replace(m(0).Value, vbCrLf, ""), "=")
	getWavInfo = s(1)
End Function

