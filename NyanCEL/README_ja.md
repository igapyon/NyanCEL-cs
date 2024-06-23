# NyanCEL

NyanCEL は Excel ブック (.xlsx) のシート内容を SQL で検索できるようにするライブラリです。

- NyanCEL-cs は C# で実装されています。
- Excel ブック (.xlsx) を与えることにより、各シートをテーブルに見立てて SQL で検索できるようにします。
- MIT ライセンスのもとで公開されています。
- NyanCEL-cs は NyanCEL プロジェクトの一部です。

## NyanQL プロジェクトとの関連性

- NyanCEL は、にゃんくるプロジェクトのフレンドプロジェクトです。
- にゃんくるプロジェクトに感銘を受けて開発された独自プロジェクトです。
- NyanCEL は、にゃんくるプロジェクトを尊敬しており、また尊重しています。

## 動作の流れ

1. Excelブック（.xlsx）ファイルを与えます
   - シート名をテーブル名として使用
   - 1行目のセル値を項目名として使用
   - 2行目のセル書式から項目データ型を導出
2. 指定した Excel ブックの内容を SQLite データベースに読み込みます
   - 2行目からデータを読み込み
   - 読み込んだデータをメモリ上の SQLite にロード
3. SELECT 文を実行します
   - 与えられた SELECT 文でデータベースを検索
   - SQLite の SQL 文法が利用可能
4. SELECT 文の検索結果を返却します
   - 行データを繰り返して検索結果を返却
   - 返却結果として、json, xml, xlsx 形式サポート

## Usage

```cs
  using (var connection = await NyanCELUtil.CreateXlsxDatabase())
  {
    var memoryStream = await NyanCELUtil.ReadBinaryFile2MemoryStream("./TestData/Book1.xlsx");
    List<NyanTableInfo> tableInfoList = await NyanXlsx2Sqlite.LoadExcelFile(
      connection, memoryStream);
    string resultString = await NyanSql2Json.Sql2Json(connection, "SELECT * FROM sqlite_master");
  }
```

## 内部的に利用している OSS

NyanCEL-cs は以下の OSS を内部的に利用しています。各 OSS の提供者に感謝します。

- ClosedXML
  - MIT
  - バージョン 0.102.2
- Microsoft.Data.Sqlite
  - MIT
  - バージョン 8.0.6
- Newtonsoft.Json
  - MIT
  - バージョン 13.0.3

## 制限

- オンメモリ RDBMS で動作するため、大量のデータを与えた場合に動作しない可能性があります
- .xlsx ファイルに対応します
- Excel のシート名やタイトル行の列名にダブルクオートを含めることはできません

## Excel => SQLite 型マッピング

内部的に、セル値の型について、Excel の書式をもとに、SQLite の型を導出します。

| Excel 書式 | SQLite 型 | 例 |
| ----------- | ---------- | --- |
| 0 | TEXT | |
| 1 | INTEGER | 0 |
| 2 | REAL | 0.00 |
| 3 | REAL | #,##0 |
| 4 | REAL | #,##0.00 |
| 9 | REAL | 0 % |
| 10 | REAL | 0.00 % |
| 11 | TEXT | 0.00E+00 |
| 12 | TEXT | # ?/? |
| 13 | TEXT | # ??/?? |
| 14 | TEXT | d / m / yyyy : Format yyyy-MM-dd |
| 15 | TEXT | d - mmm - yy : Format yyyy-MM-dd |
| 16 | TEXT | d - mmm : Format MM-dd |
| 17 | TEXT | mmm - yy : Format yyyy-MM |
| 18 | TEXT | h: mm tt |
| 19 | TEXT | h: mm: ss tt |
| 20 | TEXT | H: mm : Format HH:mm |
| 21 | TEXT | H: mm: ss : Format HH:mm:ss |
| 22 | TEXT | m / d / yyyy H: mm : Format yyyy-MM-dd HH:mm |
| 37 | INTEGER | #,##0 ;(#,##0) |
| 38 | INTEGER | #,##0 ;[Red](#,##0) |
| 39 | REAL | #,##0.00;(#,##0.00) |
| 40 | REAL | #,##0.00;[Red](#,##0.00) |
| 45 | TEXT | mm:ss : Format mm:ss |
| 46 | TEXT | [h]:mm:ss : Format HH:mm:ss |
| 47 | TEXT | mmss.0 |
| 48 | TEXT | ##0.0E+0 |
| 49 | TEXT | |

## TODO

- 動作境界系
  - エラー対応。Excel ブック読み込み失敗対応など（Excel 以外のデータを入力した場合）。ドキュメント・リムーバブル以外からのファイル読み込みエラー確認。
- テスト
  - テストケース作成
  - 同名の列が存在する xlsx ファイルの投入
  - NyanRowId を含んだ xlsx ファイルの投入
  - １行目が null な xlsx ファイルの投入
