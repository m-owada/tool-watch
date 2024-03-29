# ![WATCH](app.ico) WATCH

## 概要

アクティブウィンドウのプロセスを監視してログに記録するツールです。

## 使い方

アプリケーションを起動するとタスクトレイに常駐します。また、タスクトレイのアイコンを右クリックすることで以下の操作を行うことができます。
- 日付別に記録したログファイルの照会
- ログファイルに記録した時間の集計
- アプリケーションの終了

### 【コンパイル方法】

- 「watch.cs」と同じ場所に「build.bat」を保存して実行してください。

## 設定

- ツールを起動すると設定ファイル「config.xml」が生成されます。必要に応じて設定内容を編集してください。

### 【設定ファイル】

```
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<config>
  <interval>60000</interval>
  <directory>log</directory>
  <encoding>shift_jis</encoding>
  <history>10</history>
  <doublebuffered>false</doublebuffered>
</config>
```

### 【設定内容】

|タグ          |説明                                                        |
|--------------|------------------------------------------------------------|
|interval      |ログが記録されるインターバルをミリ秒単位に設定します。      |
|directory     |ログファイルを格納するディレクトリを設定します。            |
|encoding      |ログファイルの文字コードを設定します。                      |
|history       |照会できるログファイルの最大件数を設定します。              |
|doublebuffered|データグリッドビューがちらつく場合はtrueを設定してください。|
