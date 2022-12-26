Option Explicit
On Error Resume Next
'#┌──────────────────────────────────────
'#│  fav音声コピー v1.0.1 (2022/12/26)
'#│  favフォルダにあるテキストファイルと同名の音声ファイルをsrcフォルダからをoutフォルダにコピー
'#└──────────────────────────────────────
'#==============================================================================
'# 設定値をここに記載する。
Dim Conf
Set Conf = CreateObject("Scripting.Dictionary")
'#-- [ 基本設定 ] --------------------------------------------------------------


' 検索フォルダ
Call Conf.Add("SearchFolder", "fav")

' 元ファイルフォルダ
Call Conf.Add("SourceFolder", "src")

' 出力フォルダ
Call Conf.Add("OutputFolder", "out")


'#------------------------------------------------------------------------------
Main()
If Err.Number <> 0 Then
	WScript.Echo "エラーが発生しました。"
	WScript.Echo "エラー番号：" & Err.Number & " エラー内容：" & Err.Description
End If

WScript.Echo vbLf & "終了しました。何かキーを押すと終了します"
WScript.StdIn.ReadLine
WScript.Quit

Function Main()
	Dim ws
	Set ws = CreateObject("WScript.Shell")

	'-----------------------------------------------------------------------------
	' CSCRIPT.EXE で無い場合に切り替える
	if Ucase(Right(WScript.FullName, 11)) <> "CSCRIPT.EXE" then
		' 引数を全て読みだして、スペースが含まれていた場合は再度引き渡す為に " で囲み直しています
		Dim strParam, a
		strParam = " "
		For a = 0 to Wscript.Arguments.Count - 1
			If instr(Wscript.Arguments(a), " ") < 1 Then
				strParam = strParam & Wscript.Arguments(a) & " "
			Else
				strParam = strParam & """" & Wscript.Arguments(a) & """ "
			End If
		Next

		' CSCRIPT.EXE に引き継いで終了
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
	' カレントディレクトリの現行フォルダ参照
	Set Folder = fso.GetFolder(ws.CurrentDirectory & "\" & Conf("SearchFolder"))


	SearchFolders ws, fso, re, Folder

	Set fso = Nothing
	Set ws  = Nothing
	Main = True
'#------------------------------------------------------------------------------
End Function

Sub SearchFolders(ws, fso, re, Folder)
	Dim File, ext, subFolder, p, o, c, e, r, CodecName, SampleRate, Channels, scStr
	' Folder内のファイルを列挙する
	For Each File in Folder.Files 
		ext = LCase(fso.GetExtensionName(File.name))

		If ext="txt" Then ' 拡張子txtを判定

			p = ws.CurrentDirectory & "\" & Conf("SourceFolder") & "\" & fso.GetBaseName(File.name)
			o = ws.CurrentDirectory & "\" & Conf("OutputFolder") & "\" & fso.GetBaseName(File.name)

			If fso.FileExists(p & ".wav") Or fso.FileExists(p & ".ogg") Or fso.FileExists(p & ".mp3") Then
				If fso.FileExists(p & ".wav") Then
					fso.CopyFile p & ".wav", o & ".wav", True
					WScript.Echo File.name & "コピー(wav)"
				End If
				If fso.FileExists(p & ".ogg") Then
					fso.CopyFile p & ".ogg", o & ".ogg", True
					WScript.Echo File.name & "コピー(ogg)"
				End If
				If fso.FileExists(p & ".mp3") Then
					fso.CopyFile p & ".mp3", o & ".mp3", True
					WScript.Echo File.name & "コピー(mp3)"
				End If
			Else
					WScript.Echo File.name & "なし"
			End If


		End If
	Next

	' Foler内のフォルダを列挙する
	For Each subFolder in Folder.SubFolders 
		SearchFolders ws, fso, re, subFolder' 再帰呼び出し
	Next

End Sub

