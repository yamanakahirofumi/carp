# CARP -Check excel files of A directory by RedPen-
carpにディレクトリを指定すると、ディレクトリ内のExcel(.xlsx)ファイルの各セルを、RedPenでチェックできる

# Use (使い方)
## Requirements (前提環境)
Java11の環境を用意

## Command (コマンド)
カレントディレクトリ内をチェック
> java -jar carp-*version*.jar .

特定のフォルダをチェック
> java -jar carp-*version*.jar /home/designer/docs

他、オプションの詳細はヘルプを確認

ヘルプを表示
> java -jar carp-*version*.jar -h 

バージョン表示
> java -jar carp-*version*.jar -v

## Output Format (出力フォーマット)
-r, --result_formatで指定しない、DISPLAYまたはCSVを選択した場合、下記フォーマットで指摘が1行ごとに出力される

| 表示位置 | 情報 |
| ---: | --- |
| 1 | ファイル名 |
| 2 | シート名 |
| 3 | セルの位置 |
| 4 | エラー箇所の開始位置 |
| 5 | エラー箇所の終了位置 |
| 6 | レベル |
| 7 | メッセージ |
| 8 | エラータイプ |
| 9 | ファイルのパス |

-r, --result_formatでEXCELを選択した場合、セルのコメントに下記フォーマットで指摘が1行ごとに出力される

| 表示位置 | 情報 |
| ---: | --- |
| 1 | エラーの箇所の開始位置 |
| 2 | レベル |
| 3 | メッセージ |

# Build (ビルド方法)
Java11とMavenを動く環境を用意し下記コマンドを実行することで、実行可能なJarが作成できる

> mvn package

実行可能jarの作成先は、target/carp-*version*.jar

