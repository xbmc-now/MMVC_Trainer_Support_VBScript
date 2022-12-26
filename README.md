# 名前 / Name

MMVC_Trainer Support VBScript

# 紹介 / Features

MMVC_Trainerで機械学習するのに役立つかと思い、支援ツールをVBScriptで書いてみました。

1. katakana.vbs: テキストファイルの中からカタカナの使用数を調べてリストアップ。
2. duration.vbs: WAVファイルの長さが範囲内かを調べてリストアップ。
3. encode.vbs: 音声ファイルを機械学習用にエンコード。
4. favcopy.vbs: テキストファイルと同名の音声ファイルをコピー


# 必須条件 / Requirement

* VBScriptを実行できるWindows環境
* [ffmpeg](https://ffmpeg.org/) (ffmpeg.exe, ffprobe.exe) 

# インストール方法

音声ファイルの操作にはffmpegを使用しますので、ffmpegをダウンロードして、スクリプトファイル(.vbs)と同じフォルダにffmpeg.exeとffprobe.exeを設置してください。

# Usage

**katakana.vbsの使い方**
    katakana.vbsをダブルクリックします。「text」フォルダに入っているtxtファイルを探索してカタカナ使用数を集計した「katakana.txt」が生成されます。
    使用数が0だった場合は「●」が付きます。

**duration.vbsの使い方**
    duration.vbsをダブルクリックします。「wav」フォルダに入っているwavファイルを探索して長さを調べた「duration.txt」が生成されます。
    範囲外(0.401秒未満または、15.99秒超過)のファイルがあった場合は「●」が付きます。

**encode.vbsの使い方**
    encode.vbsをダブルクリックします。「src」フォルダに入っている音声ファイル(wav, mp3, ogg)ファイルを探索して「wav」フォルダにエンコードします。
    決められたフォーマット(24000Hz 16bit 1ch)のwavファイルだった場合は、エンコードしないでwavファイルを複製します。

**favcopy.vbsの使い方**
    favcopy.vbsをダブルクリックします。「fav」フォルダに入っているテキストファイルと同名の(wav, mp3, ogg)ファイルを「src」フォルダから探索して「out」フォルダにコピーします。
    (wav, mp3, ogg)ファイルがそれぞれ存在する場合も全てコピーします。


# Note

機械学習の勉強中です。必要なスクリプトがあれば作ろうと思います。

# 作者 / Author

* xbmc_now
* [@xbmc_now](https://twitter.com/xbmc_now)

# License
