# PowerShellSendKey

Windows上で表示中のMicrosoft Excelダイアログを検索し、Enterキーを送信するPowerShellスクリプトです。

## 動作概要

`Send-EnterToExcelDialog.ps1` はトップレベル・ウィンドウを列挙し、次の条件をすべて満たすダイアログを検索します。

- 表示中のウィンドウである
- タイトルが指定値と完全に一致する
- ウィンドウクラスが標準ダイアログの `#32770` である
- 所有プロセスが `EXCEL.exe` である

対象が見つかった場合は、最初の1件だけにEnterキーを送信し、そのまま処理を終了します。

## 必要環境

- Windows
- Windows PowerShell 5.1以上、またはPowerShell 7
- Microsoft Excel

## 使用方法

PowerShellでこのリポジトリのフォルダーへ移動します。

```powershell
Set-Location C:\Codex\PowerShellSendKey
```

### すぐに検索してEnterを送信する

```powershell
.\Send-EnterToExcelDialog.ps1
```

対象を1回検索し、見つからなければ警告を表示して終了します。

### 最大10分間待機する

```powershell
.\Send-EnterToExcelDialog.ps1 -WaitSeconds 600
```

1秒間隔で検索し、次のいずれかの時点で終了します。

- 対象を発見し、Enterキーを送信したとき
- 10分経過しても対象が見つからなかったとき

待機中は `Ctrl+C` で中断できます。

### Enterを送信せずに対象を確認する

```powershell
.\Send-EnterToExcelDialog.ps1 -ListOnly
```

10分間待機し、見つかったダイアログの情報だけを表示する場合は次のように実行します。

```powershell
.\Send-EnterToExcelDialog.ps1 -WaitSeconds 600 -ListOnly
```

### 別のタイトルを検索する

```powershell
.\Send-EnterToExcelDialog.ps1 -Title '確認' -WaitSeconds 600
```

タイトルは大文字と小文字を区別して完全一致します。

## パラメーター

| パラメーター | 既定値 | 説明 |
| --- | --- | --- |
| `Title` | `Microsoft Excel` | 検索するダイアログのタイトル |
| `ListOnly` | 無効 | Enterを送信せず、対象の情報だけを表示 |
| `WaitSeconds` | `0` | 検索を継続する最大秒数。`0`の場合は1回だけ検索 |
| `PollIntervalMilliseconds` | `1000` | 待機中の再検索間隔（ミリ秒） |

再検索間隔を500ミリ秒に変更する例：

```powershell
.\Send-EnterToExcelDialog.ps1 -WaitSeconds 600 -PollIntervalMilliseconds 500
```

## ラッパースクリプト

`Run-SendEnterWrapper.ps1` は別のWindows PowerShellプロセスを開始し、最大10分間の待機処理を実行します。

```powershell
.\Run-SendEnterWrapper.ps1
```

## ファイル構成

| ファイル | 説明 |
| --- | --- |
| `Send-EnterToExcelDialog.ps1` | Excelダイアログを検索してEnterを送信するメインスクリプト |
| `Run-SendEnterWrapper.ps1` | メインスクリプトを別プロセスで起動するラッパー |
| `msgbox.xlsm` | 動作確認用のExcelマクロ有効ブック |

## 注意事項

- Excelを管理者権限で実行している場合は、PowerShellも管理者権限で実行してください。権限レベルが異なると、Windowsの制限によりキーを送信できない場合があります。
- 同じ条件のダイアログが複数ある場合は、列挙された最初の1件だけが対象になります。
- `PostMessage` を使用してExcelのUIメッセージキューへEnterキーを登録します。物理キーボードのフォーカスは移動しません。
- 実行ポリシーによってスクリプトの実行が禁止されている場合は、組織または端末の管理方針を確認してください。

## ライセンス

このプロジェクトは[MIT License](LICENSE)のもとで公開されています。
