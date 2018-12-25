Option Explicit
Include "QRCode.vbs"

Call Example1
'Call Example2

Public Sub Example1()
    Dim sbls
    Set sbls = CreateSymbols(ERRORCORRECTION_LEVEL_M, 40, False)
    
    Call sbls.AppendText("012ABCabc")

    Call sbls.Item(0).Save24bppDIB("qrcode24bpp.bmp", 4, "#000000", "#FFFFFF")
End Sub

Public Sub Example2()
    Dim sbls
    Set sbls = CreateSymbols(ERRORCORRECTION_LEVEL_M, 1, True)
    
    Call sbls.AppendText("012ABCabc!?,./#")

    Dim sbl
    Dim i
    
    For i = 0 To sbls.Count - 1
        Call sbls.Item(i).Save1bppDIB(strDir & "\qrcode_" & CStr(i) & ".bmp", 4, "#000000", "#FFFFFF")
    Next
End Sub



Private Sub Include(ByVal strFile)
    Dim fso
    Set fso = WScript.CreateObject("Scripting.FileSystemObject") 
  
    Dim strDir
    strDir = fso.getParentFolderName(WScript.ScriptFullName)

    Dim stream
    Set stream = fso.OpenTextFile(strDir & "\" & strFile, 1)

    ExecuteGlobal stream.ReadAll() 
    Call stream.Close 
End Sub
