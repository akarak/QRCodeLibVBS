# __QRCodeLibVBS__

QRCodeLibVBSは、VBScriptで書かれたQRコード生成スクリプトです。  
JIS X 0510に基づくモデル２コードシンボルを生成します。

## 特徴

- 数字・英数字・8ビットバイト・漢字モードに対応しています
- 分割QRコードを作成可能です
- 1bppまたは24bpp BMPファイル(DIB)へ保存可能です
- 画像の配色(前景色・背景色)を指定可能です

## 使用方法

### 例１．単一シンボル(分割QRコードではない)QRコードの、最小限のコードを示します。

その他の例は、Example.bat または Example.vbs を参照してください。  

```bat
rem Command Line
CScript.exe QRCode.vbs /data:"0123ABCD" /out:"qrcode.bmp"
```

```vb
' VBScript
Public Sub Example()
    Const FORE_COLOR = "#000000"
    Const BACK_COLOR = "#FFFFFF"
    Const SCALE = 4

    Dim sbls: Set sbls = CreateSymbols(ERRORCORRECTION_M, 40, False)
    Call sbls.AppendText("012ABCabc")
    Call sbls.Item(0).Save24bppDIB("D:\qrcode.bmp", SCALE, FORE_COLOR, BACK_COLOR)
End Sub
```

### 例２．分割QRコードの作成例

型番1のデータ量を超える場合に分割し、各QRコードをビットマップファイルに保存する例を示します。

```vb
Public Sub Example()
    Const FORE_COLOR = "#000000"
    Const BACK_COLOR = "#FFFFFF"
    Const SCALE = 4

    Dim sbls: Set sbls = CreateSymbols(ERRORCORRECTION_M, 1, True)
    Call sbls.AppendText("012ABCabc!?,./#")

    Dim i
    For i = 0 To sbls.Count - 1
        ' 24bpp Bitmap File
        Call sbls.Item(i).Save24bppDIB( _
            "D:\qrcode" & CStr(i) & ".bmp", SCALE, FORE_COLOR, BACK_COLOR)

        ' 1bpp Bitmap File
'        Call sbls.Item(i).Save1bppDIB( _
'            "D:\qrcode" & CStr(i) & ".bmp", SCALE, FORE_COLOR, BACK_COLOR)
    Next
End Sub
```
