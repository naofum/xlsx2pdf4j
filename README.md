# xlsx2pdf4j

Excel（`.xlsx` / `.xlsm`）、PowerPoint（`.pptx` / `.pptm`）、Word（`.docx` / `.docm`）を Office / LibreOffice に依存せず PDF に変換する Java ライブラリです。

- 入力: `.xlsx` / `.xlsm` / `.pptx` / `.pptm` / `.docx` / `.docm`（マクロは無視）
- 出力: ページ単位のラスタ画像を埋め込んだ PDF
- 描画: Apache POI + Java2D + Apache PDFBox + XChart
- 見た目はページ画像、検索・コピー用の不可視テキスト層を付与
- Java 17 以上

## ビルド

```powershell
.\mvnw.cmd clean install
```

GitHub Actions では Windows 上の全テストと、Linux 上のコンパイル・パッケージを確認します。Maven Wrapperを同梱しているため、Maven 3.6.3未満の環境でもビルドできます。Linux / macOSでは `./mvnw clean install` を使用してください。

## 使い方

JitPack から取得できます。Maven Central には公開していません。

Maven のリポジトリ設定:

```xml
<repositories>
    <repository>
        <id>jitpack.io</id>
        <url>https://jitpack.io</url>
    </repository>
</repositories>
```

依存関係:

```xml
<dependency>
    <groupId>com.github.naofum</groupId>
    <artifactId>xlsx2pdf4j</artifactId>
    <version>v1.1.1</version>
</dependency>
```

Gradle:

```groovy
repositories {
    maven { url 'https://jitpack.io' }
}

dependencies {
    implementation 'com.github.naofum:xlsx2pdf4j:v1.1.1'
}
```

DOCX対応を含む `v1.1.1` は、タグ作成後にJitPackから取得できます。JitPackのMavenを3.9.11に固定するため、Maven Wrapperを同梱しています。

Java:

```java
import com.github.naofum.xlsx2pdf4j.api.ConversionOptions;
import com.github.naofum.xlsx2pdf4j.api.ConversionResult;
import com.github.naofum.xlsx2pdf4j.api.TextLayerMode;
import com.github.naofum.xlsx2pdf4j.api.Xlsx2Pdf;

import java.nio.file.Path;

ConversionResult result = Xlsx2Pdf.convert(
        Path.of("input.xlsx"),
        Path.of("output.pdf"),
        ConversionOptions.builder()
                .dpi(200)
                .textLayerMode(TextLayerMode.INVISIBLE)
                .build());

System.out.println(result.pageCount());
System.out.println(result.pagesBySheet());
System.out.println(result.warnings());
```

PowerPoint:

```java
import com.github.naofum.xlsx2pdf4j.api.Pptx2Pdf;
import com.github.naofum.xlsx2pdf4j.api.PptxConversionOptions;
import com.github.naofum.xlsx2pdf4j.api.PptxConversionResult;

PptxConversionResult result = Pptx2Pdf.convert(
        Path.of("slides.pptx"),
        Path.of("slides.pdf"),
        PptxConversionOptions.builder()
                .dpi(200)
                .build());

System.out.println(result.slideCount());
System.out.println(result.slideNumbers());
System.out.println(result.warnings());
```

`slides(2, 4)` または `slide(2)` で1始まりのスライド番号を指定できます。非表示スライドは、明示的に選択した場合だけ変換されます。PPTXの失敗は `Pptx2PdfException` で通知されます。

Word:

```java
import com.github.naofum.xlsx2pdf4j.api.Docx2Pdf;
import com.github.naofum.xlsx2pdf4j.api.DocxConversionOptions;
import com.github.naofum.xlsx2pdf4j.api.DocxConversionResult;

DocxConversionResult result = Docx2Pdf.convert(
        Path.of("document.docx"),
        Path.of("document.pdf"),
        DocxConversionOptions.builder()
                .dpi(200)
                .build());

System.out.println(result.pageCount());
System.out.println(result.warnings());
```

DOCXの失敗は `Docx2PdfException` で通知されます。

### Excel の主なオプション（`ConversionOptions`）

| オプション | 既定値 | 説明 |
|---|---|---|
| `dpi` | `200` | ページ画像の解像度。72〜600 |
| `textLayerMode` | `INVISIBLE` | 不可視テキスト層。`NONE` で無効化 |
| `fontScale` | `1.0` | セル・図形・ヘッダー・フッターのフォントサイズ倍率 |
| `printScale` | `0.0` | 印刷倍率の上書き。`0` はワークブックの設定を使用 |
| `strict` | `false` | 未対応要素で例外にする |
| `sheets` | すべての表示シート | 変換対象シート名 |
| `maxPages` | `200` | 最大ページ数 |
| `maxOutputBytes` | 50 MiB | 最大出力サイズ |
| `imageCompression` | `AUTO` | `PNG` / `JPEG` / `AUTO` |
| `jpegQuality` | `0.9` | JPEG の品質 |
| `pageNumberingMode` | `EXCEL_AUTOMATIC` | ヘッダー・フッターのページ番号モード |

### PowerPoint の主なオプション（`PptxConversionOptions`）

| オプション | 既定値 | 説明 |
|---|---|---|
| `dpi` | `200` | スライド画像の解像度。72〜600 |
| `textLayerMode` | `INVISIBLE` | 不可視テキスト層。`NONE` で無効化 |
| `imageCompression` | `AUTO` | `PNG` / `JPEG` / `AUTO` |
| `jpegQuality` | `0.9` | JPEG の品質 |
| `maxSlides` | `200` | 最大スライド数 |
| `maxOutputBytes` | 50 MiB | 最大出力サイズ |
| `strict` | `false` | 未対応要素で例外にする |
| `slides` | すべての表示スライド | 変換対象スライド番号（1始まり） |

### Word の主なオプション（`DocxConversionOptions`）

| オプション | 既定値 | 説明 |
|---|---|---|
| `dpi` | `200` | ページ画像の解像度。72〜600 |
| `textLayerMode` | `INVISIBLE` | 不可視テキスト層。`NONE` で無効化 |
| `imageCompression` | `AUTO` | `PNG` / `JPEG` / `AUTO` |
| `jpegQuality` | `0.9` | JPEG の品質 |
| `maxPages` | `200` | 最大ページ数 |
| `maxOutputBytes` | 50 MiB | 最大出力サイズ |
| `strict` | `false` | 未対応要素で例外にする |

`strict` モードでは、Excelは `SHAPE_UNSUPPORTED`、`IMAGE_UNSUPPORTED`、`FONT_MISSING`、`FORMULA_NO_CACHE` が例外になります。`SHAPE_APPROXIMATED` と `PAGE_LIMIT` は警告のまま処理を継続します。PowerPointは `SHAPE_UNSUPPORTED`、`IMAGE_UNSUPPORTED`、`CHART_SKIPPED`、`SMART_ART_SKIPPED` が例外になります。Wordは `SHAPE_UNSUPPORTED`、`IMAGE_UNSUPPORTED` が例外になります。

## 対応機能

### セル・テキスト

- セル値、数値書式、フォント、塗り、罫線、結合セル
- 結合セルの外周罫線、折り返し、縮小表示、文字のはみ出し
- 行高未指定の折り返しセルを内容に合わせて自動拡張（横書き・非結合セル）
- 文字回転（45° / 90° をセル中央に配置）
- セルのグラデーション塗りつぶし（線形・path、複数ストップ、主題色・インデックス色、結合セル）
- 塗り → 罫線 → テキストの3パス描画（隣接セルの塗りによる罫線欠落を防止）
- ページ境界での罫線補完
- セル内リッチテキスト
- 縦書き（`eaVert` / `wordArtVertRtl`）の upright 文字・右→左の段組み

### 印刷・ページ

- 印刷範囲、用紙、余白、拡大縮小、fit-to-page、改ページ、印刷タイトル
- 既定フォントの MDW に基づく列幅変換
- Excel の PDF 出力に合わせた行高の描画スケール
- ヘッダー、フッター、ページ番号
- ページ番号モード（Excel 自動 / シート単位のユーザー指定）

### 図形

- 四角、角丸、楕円、線、矢印、曲線矢印、吹き出し、コネクタ、フローチャート、折れ角、ホームプレート、星、シール
- 基本図形（多角形、ハート、雲、ギア、数式記号、フレーム、リボン、スクロールなど）
- 角丸・切り欠き矩形
- カスタムジオメトリ（`a:custGeom`）
- グループ図形の再帰描画
- 図形の影・グラデーション（近似）
- 図形のテーマ色と線色
- 図形内テキストの run 単位書式、段落単位配置、縦書き
- 図形テキストの余白、垂直アンカー、`normAutofit`、`anchorCtr`、`spAutoFit`
- 図形の定義済みテキスト領域
- 単純線形状・コネクタの矢印（open arrow / triangle / stealth / diamond / oval）
- 図形枠線の二重線（`a:ln cmpd="dbl"`）
- コネクタの接続点、`flipH` / `flipV`、回転、`adj`
- 図形の印刷可能領域へのクリップ

### グラフ・画像

- XChart によるグラフ描画（縦棒・横棒・折れ線・エリア・円・ドーナツ、複合グラフ、第2軸、系列色、凡例、軸・グリッド）
- グラフは縦棒・横棒・折れ線・エリア・円・ドーナツに対応
- 散布図・レーダー・バブル・株価・等高線・3D グラフと SmartArt は警告を出してスキップ
- PNG / JPEG 画像

### PowerPoint

- `.pptx` / `.pptm` のスライドをページ画像として PDF 化
- スライドマスター、レイアウト、背景、スライド番号、日付プレースホルダー
- テキスト、箇条書き、段落、オートシェイプ、コネクタ、グループ、画像、表
- グラフは XChart で上書き描画（縦棒・横棒・折れ線・エリア・円・ドーナツ）
- テーマフォントは同梱の Noto Sans JP に正規化
- 不可視テキスト層（検索・コピー用、位置は近似）
- 非表示スライドのスキップ、スライド番号・タイトルによる選択

### Word

- `.docx` / `.docm` の段落・見出し・改ページをページ単位でフロー配置
- フォント、サイズ、太字、斜体、下線、色、配置、行間、インデント
- 箇条書き・番号付き見出し（`numbering.xml` の一般的な形式）
- 基本的な表（固定列幅、罫線、セル内の段落）
- ヘッダー、フッター、ページ番号フィールド
- 単純な図形・テキストボックス（DrawingMLの`wps:wsp`）
- 不可視テキスト層（検索・コピー用、位置は近似）

### フォント・テキスト層

- Noto Sans JP / Noto Serif JP の同梱とフォールバック
- 日本語グリフを持たないフォント（Calibri 等）の自動フォールバック
- 不可視テキスト層（検索・コピー用、クリップ文字もセル全文を保持）

## 制約・未実装

- 影・グラデーションは近似
- セルの path グラデーションは放射状グラデーションで近似
- セルの線形グラデーション角度はセルの縦横比で補正しない
- EMF / WMF は対象外
- 先頭ページ・偶数ページ別ヘッダー / フッターは対象外
- `normAutofit` の自動縮小は近似
- 行高未指定の折り返しセルの自動拡張は、横書き・非結合セル・文字列（キャッシュ済み数式の文字列結果を含む）が対象。結合セル、縦書き・回転、縮小表示は対象外。リッチテキストはセル基準フォントで近似し、フォントメトリクスの差によりExcelと数pt異なる場合がある
- 図形枠線の `cmpd` は `dbl` のみ対応（`thickThin` / `thinThick` / `tri` は単線で近似）
- 円形矢印はセグメント化したリングと三角形ヘッドで近似（`adj` 未対応）
- `curvedLeftArrow` / `curvedRightArrow` の `adj` は未対応（既定形状で描画）
- リボン・スクロールは主要輪郭の近似（折り返しの陰影は単色）
- 図形テキストの段落前後スペーシング・箇条書き・インデントは未対応
- 縦書きテキストの段落単位配置・`anchorCtr` は未対応
- 未対応の図形は定義済みテキスト領域ではなく外接矩形で近似
- `vert` / `vert270` の90度回転テキストは未対応
- `curvedConnector2/3/4/5` の曲線コネクタは直線で近似
- 散布図・レーダー・バブル・株価・等高線・3D グラフは `CHART_SKIPPED` 警告でスキップ
- 横棒の積み上げ・パーセント積み上げ・第2軸・棒＋折れ線の複合グラフは未対応（`CHART_SKIPPED`）。横棒のカテゴリ順は Excel 既定（先頭カテゴリが下）に合わせる
- グラフのデータラベル・トレンドライン・誤差範囲・補助線・表示単位・データポイント単位の書式は未対応
- グラフのタイトル・軸ラベル・凡例・目盛りは画像としてのみ描画され、不可視テキスト層には含まれない
- グラフの余白・フォントサイズ・目盛り間隔・色は XChart のレイアウトで近似（Excel のピクセル単位の配置とは一致しない）
- 300 DPI の上限（200ページ / 50MB）は設定値として実装済みだが、大規模ファイルでの検証は未実施

### PowerPoint

- SmartArt、3D 図形、EMF/WMF、動画・音声、アニメーション・画面切り替えは未対応または描画できない場合がある
- テキストの太字・斜体・プレースホルダー継承は、PowerPoint の表示と完全には一致しない場合がある
- フォントは同梱の Noto に正規化するため、元のフォントと字幅・見た目が変わる場合がある
- PPTX の不可視テキスト層は図形の位置からの近似で、行単位の正確な配置ではない
- 表・グループ内のテキストは不可視テキスト層に含まれない場合がある
- 表の罫線・表スタイルは完全には再現されない場合がある
- グラフの軸ラベル・余白・凡例位置は XChart のレイアウトで近似される
- 表・画像・グループは POI の描画に依存するため、複雑な表や効果は近似になる場合がある

### Word

- ページ分割・行分割はフォントメトリクスに依存するため、Word と完全には一致しない
- 脚注・文末脚注、TOC、コメント、変更履歴、数式、SmartArt、グラフは未対応
- Word のキャンバス（`wpc:wpc`）やグループ化された図形はスキップする
- 複雑な表（可変列幅、行分割、セル結合の一部）は近似
- 浮動画像の回り込み、段組み、高度なセクション区切りは未対応
- ヘッダー/フッター内の図形・表はテキストとして近似
- Word のページ数・改ページは Word のレイアウトエンジンと完全一致しない

## テスト

```powershell
mvn test
```

### テストデータ

- `testdata/validation/attributes.xlsx`: セル・テキスト・図形・画像・印刷属性の検証用ワークブック
- `testdata/validation/unsupported-shapes.xlsx`: 外接矩形で近似されるプリセット図形のカタログ
- `testdata/validation/screen-list-ja.xlsx`: 行高未指定の折り返しセル（C6）を含む画面一覧

`screen-list-ja.xlsx` は 150 DPI で1ページです。`RowAutoFitterTest` が行高の自動拡張を、`AutoRowHeightConversionTest` が変換結果を確認します。

`attributes.xlsx` は 150 DPI で9ページ、次の5警告を出します。

- `PAGE_LIMIT`（`99 Empty`）
- `FORMULA_NO_CACHE`
- `CHART_SKIPPED`
- `IMAGE_UNSUPPORTED`
- `SHAPE_APPROXIMATED`

`unsupported-shapes.xlsx` は 100 DPI で2ページ、`SHAPE_APPROXIMATED` 警告を12件出します。

PowerPoint:

- `testdata/validation/sample1.pptx`: タイトル・本文プレースホルダー（1スライド）
- `testdata/validation/sample2.pptx`: 画像、表＋グラフ、グループ（3スライド）
- `testdata/validation/sample3.pptx`: 角丸図形・矢印コネクタ・テキストボックス（1スライド）

すべて16:9（960×540pt）です。`PptxConversionTest` が変換、ページ数、テキスト層、非表示スライドの扱いを確認します。

Word:

- `testdata/validation/sample1.docx`: タイトルページ、見出し、本文、ヘッダー/フッター、改ページ（3ページ）
- `testdata/validation/sample2.docx`: 表、見出し、本文、番号付き見出し（2ページ）

ページ数は Word 16 の表示と一致することを確認しています。`DocxConversionTest` が変換、ページ数、テキスト層を確認します。DOCXのgolden PDFは現時点では未追加です。

再生成:

```powershell
mvn -q test-compile exec:java "-Dexec.classpathScope=test" "-Dexec.mainClass=com.github.naofum.xlsx2pdf4j.ValidationWorkbookGenerator"
mvn -q test-compile exec:java "-Dexec.classpathScope=test" "-Dexec.mainClass=com.github.naofum.xlsx2pdf4j.UnsupportedShapesWorkbookGenerator"
```

### golden PDF

`testdata/golden/attributes.pdf` と `testdata/golden/unsupported-shapes.pdf` は、Excel 16（Windows）で各ワークブックの全シートを PDF 出力した参照 PDF です。

`testdata/golden/sample1.pdf`、`sample2.pdf`、`sample3.pdf` は、PowerPoint 16（Windows）で各プレゼンテーションを PDF 出力した参照 PDF です。メタデータはコミット前に除去しています。

```powershell
mvn -Dtest=GoldenPdfComparisonTest test
mvn -Dtest=PptxGoldenPdfComparisonTest test
```

- 出力先: `target/golden-diff/attributes/`, `target/golden-diff/unsupported-shapes/`, `target/golden-diff/pptx-sample1/`, `target/golden-diff/pptx-sample2/`, `target/golden-diff/pptx-sample3/`
- 各ページ: `page-NN-diff.png`
  - 左: 実装の出力
  - 中央: Excel / PowerPoint の参照 PDF
  - 右: 差分（4倍に増幅、赤いほど差が大きい）
- `summary.txt`: ページごとの平均画像差と、差の大きいページの一覧

差分画像が不要な場合:

```powershell
mvn -Dtest=GoldenPdfComparisonTest "-Dxlsx2pdf4j.diffImages=false" test
mvn -Dtest=PptxGoldenPdfComparisonTest "-Dxlsx2pdf4j.diffImages=false" test
```

golden は Excel / PowerPoint のバージョン、プリンタードライバー、フォント環境に依存します。再生成した場合はページ数と差分を確認してください。

## ライセンス

Apache License, Version 2.0 です。詳細は [LICENSE](LICENSE) を参照してください。

同梱の Noto Sans JP / Noto Serif JP は SIL Open Font License 1.1 で提供されています。ライセンス全文は `src/main/resources/fonts/` に含まれています。
