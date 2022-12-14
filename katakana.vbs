Option Explicit
On Error Resume Next
'#┌──────────────────────────────────────
'#│  カタカナカウント v1.0.1 (2022/12/08)
'#└──────────────────────────────────────
'#==============================================================================
'# 設定値をここに記載する。
Dim Conf
Set Conf = CreateObject("Scripting.Dictionary")
'#-- [ 基本設定 ] --------------------------------------------------------------
' カタカナリスト
Call Conf.Add("KanaList", Array( _
	"ア", "イ", "ウ", "エ", "オ", _
	"カ", "キ", "ク", "ケ", "コ", _
	"サ", "シ", "ス", "セ", "ソ", _
	"タ", "チ", "ツ", "テ", "ト", _
	"ナ", "ニ", "ヌ", "ネ", "ノ", _
	"ハ", "ヒ", "フ", "ヘ", "ホ", _
	"マ", "ミ", "ム", "メ", "モ", _
	"ヤ", "ユ", "ヨ", _
	"ラ", "リ", "ル", "レ", "ロ", _
	"ワ", "ヲ", "ン", _
	"ガ", "ギ", "グ", "ゲ", "ゴ", _
	"ザ", "ジ", "ズ", "ゼ", "ゾ", _
	"ダ", "ヂ", "ヅ", "デ", "ド", _
	"バ", "ビ", "ブ", "ベ", "ボ", _
	"パ", "ピ", "プ", "ペ", "ポ", _
	"キャ", "キュ", "キョ", _
	"シャ", "シュ", "ショ", _
	"チャ", "チュ", "チョ", _
	"ニャ", "ニュ", "ニョ", _
	"ヒャ", "ヒュ", "ヒョ", _
	"ミャ", "ミュ", "ミョ", _
	"リャ", "リュ", "リョ", _
	"ギャ", "ギュ", "ギョ", _
	"ジャ", "ジュ", "ジョ", _
	"ビャ", "ビュ", "ビョ", _
	"ピャ", "ピュ", "ピョ", _
	"ウィ", "ウェ", "ウォ", _
	"ヴァ", "ヴィ", "ヴェ", "ヴォ", _
	"シェ", "ジェ", _
	"チェ", "ツァ", "ツィ", "ツェ", "ツォ", _
	"ティ", "ディ", "デュ", "トゥ", _
	"ファ", "フィ", "フェ", "フォ", "フュ" _
))

' 検索フォルダ
Call Conf.Add("SearchFolder", "text")

' 検索結果テキスト
Call Conf.Add("SearchResult", "katakana.txt")

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

	Dim ado
	Set ado = CreateObject("ADODB.Stream")
	ado.Type = 2' 読み込むファイルの形式（テキストファイル）
	ado.Charset = "UTF-8"' 文字コード
	ado.LineSeparator = -1' 改行コード（CRLF）

	Dim r, p
	Set r = New RegExp
	r.Global = True

	Dim Folder
	' カレントディレクトリの現行フォルダ参照
	Set Folder = fso.GetFolder(ws.CurrentDirectory & "\" & Conf("SearchFolder"))

	Dim wrt
	Set wrt = fso.OpenTextFile(fso.BuildPath(ws.CurrentDirectory, Conf("SearchResult")), 2, True)

	Dim k, c
	For Each k in Conf("KanaList")
		c = 0
		r.Pattern = k

		SearchFolders fso, ado, Folder, r, c

		If c = 0 Then
			wrt.WriteLine "●" & k & ":" & c
			WScript.Echo "●" & k & ":" & c
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
	' Folder内のファイルを列挙する
	For Each File in Folder.Files 

		If LCase(fso.GetExtensionName(File.name))="txt" Then ' 拡張子txtを判定
			p = Folder & "\" & File.name

			ado.Open' ストリームをオープン
			ado.LoadFromFile p

			Dim b
			b = ado.ReadText(-1)

			' 検索
			Set m = r.Execute(b)
			c = c + m.Count

			ado.Close
		End If
	Next

	' Foler内のフォルダを列挙する
	For Each subFolder in Folder.SubFolders 
		SearchFolders fso, ado, subFolder, r, c' 再帰呼び出し
	Next

End Sub
