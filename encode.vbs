Option Explicit
On Error Resume Next
'#┌──────────────────────────────────────
'#│  wavエンコード v1.0.1 (2022/12/08)
'#│  srcフォルダにある音声ファイルをwavフォルダにエンコード
'#└──────────────────────────────────────
'#==============================================================================
'# 設定値をここに記載する。
Dim Conf
Set Conf = CreateObject("Scripting.Dictionary")
'#-- [ 基本設定 ] --------------------------------------------------------------


' 検索フォルダ
Call Conf.Add("SearchFolder", "src")

' 出力フォルダ
Call Conf.Add("OutputFolder", "wav")

' 検索結果テキスト
Call Conf.Add("SearchResult", "encode.txt")

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
	' Folder内のファイルを列挙する
	For Each File in Folder.Files 
		ext = LCase(fso.GetExtensionName(File.name))

		If ext="wav" Or ext="ogg" Or ext="mp3" Then ' 拡張子wav, ogg, mp3を判定

			p = Folder & "\" & File.name
			c = ws.CurrentDirectory & "\" & "ffprobe.exe -loglevel quiet -show_streams " & p
			Set e = ws.Exec(c)
			r = e.StdOut.ReadAll

			CodecName  = getWavInfo(re, r, "codec_name")
			SampleRate = getWavInfo(re, r, "sample_rate")
			Channels   = getWavInfo(re, r, "channels")

			o = ws.CurrentDirectory & "\" & Conf("OutputFolder") & "\" & fso.GetBaseName(File.name) & ".wav"
			If ext = "wav" AND CodecName = "pcm_s16le" AND SampleRate = "24000" AND Channels = "1" Then
				' 24000Hz 16bit 1chのwavファイルであれば、エンコードせずコピー
				fso.CopyFile p, o, True
				WScript.Echo File.name & "コピー"
				wrt.WriteLine File.name & "コピー"
			Else
				c = ws.CurrentDirectory & "\" & "ffmpeg.exe -y -i " & p & " -acodec pcm_s16le -ar 24000 -ac 1 " & o
				Set e = ws.Exec(c)
				'r = e.StdErr.ReadAll
				WScript.Echo File.name & "エンコード"
				wrt.WriteLine File.name & "エンコード"
			End If
		End If
	Next

	' Foler内のフォルダを列挙する
	For Each subFolder in Folder.SubFolders 
		SearchFolders ws, fso, re, wrt, subFolder' 再帰呼び出し
	Next

End Sub

Function getWavInfo(re, str, item)
	Dim m, s
	re.Pattern = item & "=.+\n"
	Set m = re.Execute(str)
	s = Split(Replace(m(0).Value, vbCrLf, ""), "=")
	getWavInfo = s(1)
End Function

