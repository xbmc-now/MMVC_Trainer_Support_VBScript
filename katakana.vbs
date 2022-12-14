Option Explicit
On Error Resume Next
'#������������������������������������������������������������������������������
'#��  �J�^�J�i�J�E���g v1.0.1 (2022/12/08)
'#������������������������������������������������������������������������������
'#==============================================================================
'# �ݒ�l�������ɋL�ڂ���B
Dim Conf
Set Conf = CreateObject("Scripting.Dictionary")
'#-- [ ��{�ݒ� ] --------------------------------------------------------------
' �J�^�J�i���X�g
Call Conf.Add("KanaList", Array( _
	"�A", "�C", "�E", "�G", "�I", _
	"�J", "�L", "�N", "�P", "�R", _
	"�T", "�V", "�X", "�Z", "�\", _
	"�^", "�`", "�c", "�e", "�g", _
	"�i", "�j", "�k", "�l", "�m", _
	"�n", "�q", "�t", "�w", "�z", _
	"�}", "�~", "��", "��", "��", _
	"��", "��", "��", _
	"��", "��", "��", "��", "��", _
	"��", "��", "��", _
	"�K", "�M", "�O", "�Q", "�S", _
	"�U", "�W", "�Y", "�[", "�]", _
	"�_", "�a", "�d", "�f", "�h", _
	"�o", "�r", "�u", "�x", "�{", _
	"�p", "�s", "�v", "�y", "�|", _
	"�L��", "�L��", "�L��", _
	"�V��", "�V��", "�V��", _
	"�`��", "�`��", "�`��", _
	"�j��", "�j��", "�j��", _
	"�q��", "�q��", "�q��", _
	"�~��", "�~��", "�~��", _
	"����", "����", "����", _
	"�M��", "�M��", "�M��", _
	"�W��", "�W��", "�W��", _
	"�r��", "�r��", "�r��", _
	"�s��", "�s��", "�s��", _
	"�E�B", "�E�F", "�E�H", _
	"���@", "���B", "���F", "���H", _
	"�V�F", "�W�F", _
	"�`�F", "�c�@", "�c�B", "�c�F", "�c�H", _
	"�e�B", "�f�B", "�f��", "�g�D", _
	"�t�@", "�t�B", "�t�F", "�t�H", "�t��" _
))

' �����t�H���_
Call Conf.Add("SearchFolder", "text")

' �������ʃe�L�X�g
Call Conf.Add("SearchResult", "katakana.txt")

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

	Dim ado
	Set ado = CreateObject("ADODB.Stream")
	ado.Type = 2' �ǂݍ��ރt�@�C���̌`���i�e�L�X�g�t�@�C���j
	ado.Charset = "UTF-8"' �����R�[�h
	ado.LineSeparator = -1' ���s�R�[�h�iCRLF�j

	Dim r, p
	Set r = New RegExp
	r.Global = True

	Dim Folder
	' �J�����g�f�B���N�g���̌��s�t�H���_�Q��
	Set Folder = fso.GetFolder(ws.CurrentDirectory & "\" & Conf("SearchFolder"))

	Dim wrt
	Set wrt = fso.OpenTextFile(fso.BuildPath(ws.CurrentDirectory, Conf("SearchResult")), 2, True)

	Dim k, c
	For Each k in Conf("KanaList")
		c = 0
		r.Pattern = k

		SearchFolders fso, ado, Folder, r, c

		If c = 0 Then
			wrt.WriteLine "��" & k & ":" & c
			WScript.Echo "��" & k & ":" & c
		Else
			wrt.WriteLine k & ":" & c
			WScript.Echo k & ":" & c
		End If
	Next
	wrt.Close

	Set ado = Nothing
	Set fso = Nothing
	Set ws  = Nothing
	Main = True
'#------------------------------------------------------------------------------
End Function

Sub SearchFolders(fso, ado, Folder, r, c )
	Dim File, subFolder, m, p
	' Folder���̃t�@�C����񋓂���
	For Each File in Folder.Files 

		If LCase(fso.GetExtensionName(File.name))="txt" Then ' �g���qtxt�𔻒�
			p = Folder & "\" & File.name

			ado.Open' �X�g���[�����I�[�v��
			ado.LoadFromFile p

			Dim b
			b = ado.ReadText(-1)

			' ����
			Set m = r.Execute(b)
			c = c + m.Count

			ado.Close
		End If
	Next

	' Foler���̃t�H���_��񋓂���
	For Each subFolder in Folder.SubFolders 
		SearchFolders fso, ado, subFolder, r, c' �ċA�Ăяo��
	Next

End Sub
