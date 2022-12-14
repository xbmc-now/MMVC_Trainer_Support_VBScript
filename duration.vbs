Option Explicit
On Error Resume Next
'#┌──────────────────────────────────────
'#│  wavの長さチェック v1.0.1 (2022/12/08)
'#│  0.401秒以上または、15.99秒以下
'#└──────────────────────────────────────
'#==============================================================================
'# 設定値をここに記載する。
Dim Conf
Set Conf = CreateObject("Scripting.Dictionary")
'#-- [ 基本設定 ] --------------------------------------------------------------


' 検索フォルダ
Call Conf.Add("SearchFolder", "wav")

' 検索結果テキスト
Call Conf.Add("SearchResult", "duration.txt")

'#------------------------------------------------------------------------------
Main()
If Err.Number <> 0 Then
	WScript.Echo "エラーが発生しました。"
	WScript.Echo "エラー番号：" & Err.Number & " " & "エラー内容：" & Err.Description
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

	Dim Folder
	' カレントディレクトリの現行フォルダ参照
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
	' Folder内のファイルを列挙する
	For Each File in Folder.Files 

		If LCase(fso.GetExtensionName(File.name))="wav" Then ' 拡張子wavを判定
			p = Folder & "\" & File.name
			c = ws.CurrentDirectory & "\" & "ffprobe.exe -loglevel quiet -show_streams " & p
			Set e = ws.Exec("cmd /c " & c & " 2>&1 | findstr -i duration=")
			r = Split(Replace(e.StdOut.ReadAll, vbCrLf, ""), "=")
			s = CDbl(r(1))

			' 0.401秒以上または、15.99秒以下
			If s < 0.401 Or s > 15.99 Then
				wrt.WriteLine "●" & File.name & ":" & s
				WScript.Echo "●" & File.name & ":" & s
			Else
				wrt.WriteLine File.name & ":" & s
				WScript.Echo File.name & ":" & s
			End If
		End If
	Next

	' Foler内のフォルダを列挙する
	For Each subFolder in Folder.SubFolders 
		SearchFolders ws, fso, wrt, subFolder' 再帰呼び出し
	Next

End Sub
