## 概要
時間をDiscordとExcelの両方に保存できるタイムスタンプです。


## 使い方
exeファイル内の.env.exampleを.envにrenameします。

ファイルにDiscordの投稿したいチャンネルのWebHookURLを保存します。

main.exeファイルを開き、startでDiscordに時間が投稿され、同時にexcel
のTimeSheetに時間が記録されます。

閉じるときはエクセルではなく、exeファイルを閉じてください。

なお、Record単体を開いて、直接編集することも可能です。