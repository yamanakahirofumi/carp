# CARP -Check excel files of A directory by RedPen-
carpにディレクトリを指定すると、ディレクトリ内のExcel(.xlsx)ファイルの各セルを、RedPenでチェックできる

# Use (使い方)
## Requirements (前提環境)
Java11の環境を用意する。

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

# Build (ビルド方法)
Java11とMavenを動く環境を用意し下記コマンドを実行することで、実行可能なJarが作成できる

> mvn package

実行可能jarの作成先は、target/carp-*version*.jar

