Option Explicit
Include "QRCode.vbs"

Const FORE_COLOR = "#000000"
Const BACK_COLOR = "#FFFFFF"
Const SCALE = 4


Call Example1
'Call Example2


Public Sub Example1()
    Dim sbls: Set sbls = CreateSymbols(ERRORCORRECTION_M, 40, False)
    Call sbls.AppendText("012ABCabc")
    ' 24bpp Bitmap File
    Call sbls.Item(0).Save24bppDIB("qrcode24bpp.bmp", SCALE, FORE_COLOR, BACK_COLOR)
    ' 1bpp Bitmap File
'    Call sbls.Item(0).Save1bppDIB("qrcode1bpp.bmp", SCALE, FORE_COLOR, BACK_COLOR)
End Sub


Public Sub Example2()
    Dim sbls: Set sbls = CreateSymbols(ERRORCORRECTION_M, 1, True)
    Call sbls.AppendText("012ABCabc!?,./#")

    Dim fName
    Dim sbl, i
    For i = 0 To sbls.Count - 1
        fName = "qrcode" & CStr(i) & ".bmp"
        ' 24bpp Bitmap File
        Call sbls.Item(i).Save24bppDIB(fName, SCALE, FORE_COLOR, BACK_COLOR)
        ' 1bpp Bitmap File
'        Call sbls.Item(i).Save1bppDIB(fName, SCALE, FORE_COLOR, BACK_COLOR)
    Next
End Sub


Private Sub Include(ByVal strFile)
    Dim fso: Set fso = WScript.CreateObject("Scripting.FileSystemObject") 
    Dim strDir: strDir = fso.getParentFolderName(WScript.ScriptFullName)
    Dim stream: Set stream = fso.OpenTextFile(strDir & "\" & strFile, 1)

    ExecuteGlobal stream.ReadAll() 
    Call stream.Close 
End Sub
