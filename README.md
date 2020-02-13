# amazhist

amazhist は，Amazon の買い物履歴情報を取得し，Excel 形式で出力する Ruby スクリプトです．

次のような特徴があります．

- 買い物リストへのサムネイル画像の埋め込み
- 各種グラフの自動生成

# 準備

次のコマンドを実行して必要なモジュールをインストールします．

`bundle install`

環境変数 `amazon_id` と `amazon_pass` に Amazon のログイン情報をセットしておきます．

    export amazon_id="Amazon のログイン ID"
    export amazon_pass="Amazon のパスワード"

Windows の PowerShell の場合は次のようにします．

    $env:amazon_id="Amazon のログイン ID"
    $env:amazon_pass="Amazon のパスワード"


# 実行

## 履歴データの取得

次のコマンドを実行すると，Amazon にログインして購入履歴情報の取得を行います．

`./amazhist.rb -j amazhist.json -t img`

### オプションの意味

`-j` ファイル名
: 履歴データを保存するファイルの名前を指定します．履歴データは  JSON 形式で保存されます．

`-t` ディレクトリ名
: サムネイル画像を保存するディレクトリの名前を指定します．


## Excel ファイルの精製

次のコマンドを実行すると，先ほど収集したデータから Excel ファイルを生成します．
Win32OLE を使っている為，Windows 上の Ruby でのみ実行できます．

`./amazexcel.rb -j amazhist.json -t img -o amazhist.xlsx`

### オプションの意味

`-j` ファイル名
: 履歴データを保存するファイルの名前を指定します．履歴データは  JSON 形式で保存されます．

`-t` ディレクトリ名
: サムネイル画像を保存するディレクトリの名前を指定します．

`-o` Excelファイル名
: 生成する Excel ファイルの名前を指定します．


# 参考方法

出力サンプル等については「[Amazon の買い物履歴情報のビジュアル化 【2018年 Excel 版】](https://rabbit-note.com/2018/01/04/amazon-purchase-history-report-2018/)」にて紹介しています．

