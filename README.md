# SealAddIn

Excel用 VSTOアドイン。画像ライブラリからの画像配置・図形のサイズ/位置編集・電子印影(スタンプ)の生成と押印・押印履歴の確認ができます。

## 主な機能

- **画像ライブラリ**: 個人用/共有フォルダに画像を登録し、選択セルに配置
- **サイズ・位置編集**: シート上の図形の位置・サイズをmm/pt/px単位で編集
- **電子押印(スタンプ生成)**: 印影(丸印/角印)を生成し、選択セルに押印
- **押印履歴**: 直近の押印記録(押印者・日時・セル・印影テキスト)を一覧表示

## 動作環境

- Windows + Microsoft Excel(デスクトップ版)
- .NET Framework 4.7.2
- Visual Studio Tools for Office (VSTO) ランタイム

## インストール

1. [Releases](../../releases) から最新のzipをダウンロードし、展開します。
2. `setup.exe` を実行してインストールします。
3. Excelを起動すると、リボンに「電子印影」タブが表示されます。

## 開発者向けセットアップ

このリポジトリには署名用の秘密鍵(`.pfx`)は含まれていません。ローカルでビルドする場合は、以下のPowerShellコマンドで開発用の自己署名証明書を作成し、プロジェクト直下(`SealAddIn/SealAddIn_TemporaryKey.pfx`)に配置してください。

```powershell
$cert = New-SelfSignedCertificate -Type Custom -Subject "CN=SealAddIn" -KeyUsage DigitalSignature `
  -FriendlyName "SealAddIn Temporary Key" -CertStoreLocation "Cert:\CurrentUser\My" `
  -TextExtension @("2.5.29.37={text}1.3.6.1.5.5.7.3.3","2.5.29.19={text}Subject Type:End Entity")
$bytes = $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pfx, "")
[System.IO.File]::WriteAllBytes("SealAddIn/SealAddIn_TemporaryKey.pfx", $bytes)
```

作成した証明書のサムプリント(`$cert.Thumbprint`)を `SealAddIn.csproj` の `ManifestCertificateThumbprint` に設定してください。

### ビルド

```powershell
msbuild SealAddIn.sln /p:Configuration=Debug /p:Platform="Any CPU"
```

### オプション設定

共有ライブラリ・押印ログの保存先フォルダーは、リボンの「オプション設定」ボタンから設定してください(未設定の場合、共有ライブラリ/押印履歴機能は利用できません)。

## ライセンス

社内利用を想定しています。
