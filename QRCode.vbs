Option Explicit

Public Const MIN_VERSION = 1
Public Const MAX_VERSION = 40

Public Const ERRORCORRECTION_LEVEL_L = 0
Public Const ERRORCORRECTION_LEVEL_M = 1
Public Const ERRORCORRECTION_LEVEL_Q = 2
Public Const ERRORCORRECTION_LEVEL_H = 3

Private Const ENCODINGMODE_UNKNOWN           = 0
Private Const ENCODINGMODE_NUMERIC           = 1
Private Const ENCODINGMODE_ALPHA_NUMERIC     = 2
Private Const ENCODINGMODE_EIGHT_BIT_BYTE    = 3
Private Const ENCODINGMODE_KANJI             = 4

Private Const MODEINDICATOR_LENGTH = 4
Private Const MODEINDICATOR_TERMINATOR_VALUE           = &H0
Private Const MODEINDICATOR_NUMERIC_VALUE              = &H1
Private Const MODEINDICATOR_ALPAHNUMERIC_VALUE         = &H2
Private Const MODEINDICATOR_STRUCTURED_APPEND_VALUE    = &H3
Private Const MODEINDICATOR_BYTE_VALUE                 = &H4
Private Const MODEINDICATOR_KANJI_VALUE                = &H8

Private Const SYMBOLSEQUENCEINDICATOR_POSITION_LENGTH     = 4
Private Const SYMBOLSEQUENCEINDICATOR_TOTAL_NUMBER_LENGTH = 4

Private Const STRUCTUREDAPPEND_PARITY_DATA_LENGTH = 8
Private Const STRUCTUREDAPPEND_HEADER_LENGTH      = 20

Private Const adTypeBinary = 1
Private Const adTypeText = 2
Private Const adSaveCreateOverWrite = 2

Private AlignmentPattern:     Set AlignmentPattern = New AlignmentPattern_
Private CharCountIndicator:   Set CharCountIndicator = New CharCountIndicator_
Private Codeword:             Set Codeword = New Codeword_
Private DataCodeword:         Set DataCodeword = New DataCodeword_
Private FinderPattern:        Set FinderPattern = New FinderPattern_
Private FormatInfo:           Set FormatInfo = New FormatInfo_
Private GaloisField256:       Set GaloisField256 = New GaloisField256_
Private GeneratorPolynomials: Set GeneratorPolynomials = New GeneratorPolynomials_
Private Masking:              Set Masking = New Masking_
Private MaskingPenaltyScore:  Set MaskingPenaltyScore = New MaskingPenaltyScore_
Private Module:               Set Module = New Module_
Private QuietZone:            Set QuietZone = New QuietZone_
Private RemainderBit:         Set RemainderBit = New RemainderBit_
Private RSBlock:              Set RSBlock = New RSBlock_
Private Separator:            Set Separator = New Separator_
Private TimingPattern:        Set TimingPattern = New TimingPattern_
Private VersionInfo:          Set VersionInfo = New VersionInfo_


Call Main(WScript.Arguments)


Public Function ToRGB(ByVal arg)
    Dim re
    Set re = CreateObject("VBScript.RegExp")

    re.Pattern = "^#[0-9A-Fa-f]{6}$"
    If Not re.Test(arg) Then Call Err.Raise(5)

    Dim ret
    ret = RGB(CInt("&h" & Mid(arg, 2, 2)), _
              CInt("&h" & Mid(arg, 4, 2)), _
              CInt("&h" & Mid(arg, 6, 2)))

    ToRGB = ret
End Function

Public Function Build1bppDIB( _
  ByRef bitmapData, ByVal pictWidth, ByVal pictHeight, ByVal foreRGB, ByVal backRGB)
    Dim bfh
    Set bfh = New BITMAPFILEHEADER
    With bfh
        .bfType = &H4D42
        .bfSize = 62 + bitmapData.Size
        .bfReserved1 = 0
        .bfReserved2 = 0
        .bfOffBits = 62
    End With

    Dim bih
    Set bih = New BITMAPINFOHEADER
    With bih
        .biSize = 40
        .biWidth = pictWidth
        .biHeight = pictHeight
        .biPlanes = 1
        .biBitCount = 1
        .biCompression = 0
        .biSizeImage = 0
        .biXPelsPerMeter = 3780
        .biYPelsPerMeter = 3780
        .biClrUsed = 0
        .biClrImportant = 0
    End With

    Dim palette(1)
    Set palette(0) = New RGBQUAD
    Set palette(1) = New RGBQUAD

    With palette(0)
        .rgbBlue = (foreRGB And &HFF0000) \ 2 ^ 16
        .rgbGreen = (foreRGB And &HFF00&) \ 2 ^ 8
        .rgbRed = foreRGB And &HFF&
        .rgbReserved = 0
    End With

    With palette(1)
        .rgbBlue = (backRGB And &HFF0000) \ 2 ^ 16
        .rgbGreen = (backRGB And &HFF00&) \ 2 ^ 8
        .rgbRed = backRGB And &HFF&
        .rgbReserved = 0
    End With

    Dim ret
    Set ret = New BinaryWriter

    With bfh
        Call ret.Append(.bfType)
        Call ret.Append(.bfSize)
        Call ret.Append(.bfReserved1)
        Call ret.Append(.bfReserved2)
        Call ret.Append(.bfOffBits)        
    End With

    With bih
        Call ret.Append(.biSize)
        Call ret.Append(.biWidth)
        Call ret.Append(.biHeight)
        Call ret.Append(.biPlanes)
        Call ret.Append(.biBitCount)
        Call ret.Append(.biCompression)
        Call ret.Append(.biSizeImage)
        Call ret.Append(.biXPelsPerMeter)
        Call ret.Append(.biYPelsPerMeter)
        Call ret.Append(.biClrUsed)
        Call ret.Append(.biClrImportant)
    End With

    With palette(0)
        Call ret.Append(.rgbBlue)
        Call ret.Append(.rgbGreen)
        Call ret.Append(.rgbRed)
        Call ret.Append(.rgbReserved)
    End With

    With palette(1)
        Call ret.Append(.rgbBlue)
        Call ret.Append(.rgbGreen)
        Call ret.Append(.rgbRed)
        Call ret.Append(.rgbReserved)
    End With

    Call bitmapData.CopyTo(ret)

    Set Build1bppDIB = ret
End Function

Public Function Build24bppDIB( _
  ByRef bitmapData, ByVal pictWidth, ByVal pictHeight)
    Dim bfh
    Set bfh = New BITMAPFILEHEADER

    With bfh
        .bfType = &H4D42
        .bfSize = 54 + bitmapData.Size
        .bfReserved1 = 0
        .bfReserved2 = 0
        .bfOffBits = 54
    End With

    Dim bih
    Set bih = New BITMAPINFOHEADER

    With bih
        .biSize = 40
        .biWidth = pictWidth
        .biHeight = pictHeight
        .biPlanes = 1
        .biBitCount = 24
        .biCompression = 0
        .biSizeImage = 0
        .biXPelsPerMeter = 3780
        .biYPelsPerMeter = 3780
        .biClrUsed = 0
        .biClrImportant = 0
    End With

    Dim ret
    Set ret = New BinaryWriter

    With bfh
        Call ret.Append(.bfType)
        Call ret.Append(.bfSize)
        Call ret.Append(.bfReserved1)
        Call ret.Append(.bfReserved2)
        Call ret.Append(.bfOffBits)
    End With

    With bih
        Call ret.Append(.biSize)
        Call ret.Append(.biWidth)
        Call ret.Append(.biHeight)
        Call ret.Append(.biPlanes)
        Call ret.Append(.biBitCount)
        Call ret.Append(.biCompression)
        Call ret.Append(.biSizeImage)
        Call ret.Append(.biXPelsPerMeter)
        Call ret.Append(.biYPelsPerMeter)
        Call ret.Append(.biClrUsed)
        Call ret.Append(.biClrImportant)
    End With

    Call bitmapData.CopyTo(ret)

    Set Build24bppDIB = ret
End Function

Public Function CreateSymbols( _
  ByVal ecLevel, ByVal maxVer, ByVal allowStructuredAppend)
    Dim ret
    Set ret = New Symbols
    Call ret.Init(ecLevel, maxVer, allowStructuredAppend)

    Set CreateSymbols = ret
End Function

Public Function CreateEncoder(ByVal encMode)
    Dim ret

    Select Case encMode
        Case ENCODINGMODE_NUMERIC
            Set ret = New NumericEncoder
        Case ENCODINGMODE_ALPHA_NUMERIC
            Set ret = New AlphanumericEncoder
        Case ENCODINGMODE_EIGHT_BIT_BYTE
            Set ret = New ByteEncoder
        Case ENCODINGMODE_KANJI
            Set ret = New KanjiEncoder
        Case Else
            Call Err.Raise(5)
    End Select

    Set CreateEncoder = ret
End Function


Class AlignmentPattern_
    Private m_centerPosArrays(40)

    Private Sub Class_Initialize()
        m_centerPosArrays(2) = Array(6, 18)
        m_centerPosArrays(3) = Array(6, 22)
        m_centerPosArrays(4) = Array(6, 26)
        m_centerPosArrays(5) = Array(6, 30)
        m_centerPosArrays(6) = Array(6, 34)
        m_centerPosArrays(7) = Array(6, 22, 38)
        m_centerPosArrays(8) = Array(6, 24, 42)
        m_centerPosArrays(9) = Array(6, 26, 46)
        m_centerPosArrays(10) = Array(6, 28, 50)
        m_centerPosArrays(11) = Array(6, 30, 54)
        m_centerPosArrays(12) = Array(6, 32, 58)
        m_centerPosArrays(13) = Array(6, 34, 62)
        m_centerPosArrays(14) = Array(6, 26, 46, 66)
        m_centerPosArrays(15) = Array(6, 26, 48, 70)
        m_centerPosArrays(16) = Array(6, 26, 50, 74)
        m_centerPosArrays(17) = Array(6, 30, 54, 78)
        m_centerPosArrays(18) = Array(6, 30, 56, 82)
        m_centerPosArrays(19) = Array(6, 30, 58, 86)
        m_centerPosArrays(20) = Array(6, 34, 62, 90)
        m_centerPosArrays(21) = Array(6, 28, 50, 72, 94)
        m_centerPosArrays(22) = Array(6, 26, 50, 74, 98)
        m_centerPosArrays(23) = Array(6, 30, 54, 78, 102)
        m_centerPosArrays(24) = Array(6, 28, 54, 80, 106)
        m_centerPosArrays(25) = Array(6, 32, 58, 84, 110)
        m_centerPosArrays(26) = Array(6, 30, 58, 86, 114)
        m_centerPosArrays(27) = Array(6, 34, 62, 90, 118)
        m_centerPosArrays(28) = Array(6, 26, 50, 74, 98, 122)
        m_centerPosArrays(29) = Array(6, 30, 54, 78, 102, 126)
        m_centerPosArrays(30) = Array(6, 26, 52, 78, 104, 130)
        m_centerPosArrays(31) = Array(6, 30, 56, 82, 108, 134)
        m_centerPosArrays(32) = Array(6, 34, 60, 86, 112, 138)
        m_centerPosArrays(33) = Array(6, 30, 58, 86, 114, 142)
        m_centerPosArrays(34) = Array(6, 34, 62, 90, 118, 146)
        m_centerPosArrays(35) = Array(6, 30, 54, 78, 102, 126, 150)
        m_centerPosArrays(36) = Array(6, 24, 50, 76, 102, 128, 154)
        m_centerPosArrays(37) = Array(6, 28, 54, 80, 106, 132, 158)
        m_centerPosArrays(38) = Array(6, 32, 58, 84, 110, 136, 162)
        m_centerPosArrays(39) = Array(6, 26, 54, 82, 110, 138, 166)
        m_centerPosArrays(40) = Array(6, 30, 58, 86, 114, 142, 170)
    End Sub

    Public Sub Place(ByRef moduleMatrix(), ByVal ver)
        Dim centerArray
        centerArray = m_centerPosArrays(ver)

        Dim maxIndex
        maxIndex = UBound(centerArray)

        Dim i, j
        Dim r, c

        For i = 0 To maxIndex
            r = centerArray(i)

            For j = 0 To maxIndex
                c = centerArray(j)

                If (i = 0 And j = 0 Or _
                    i = 0 And j = maxIndex Or _
                    i = maxIndex And j = 0) = False Then

                    moduleMatrix(r - 2)(c - 2) = 2
                    moduleMatrix(r - 2)(c - 1) = 2
                    moduleMatrix(r - 2)(c + 0) = 2
                    moduleMatrix(r - 2)(c + 1) = 2
                    moduleMatrix(r - 2)(c + 2) = 2

                    moduleMatrix(r - 1)(c - 2) = 2
                    moduleMatrix(r - 1)(c - 1) = -2
                    moduleMatrix(r - 1)(c + 0) = -2
                    moduleMatrix(r - 1)(c + 1) = -2
                    moduleMatrix(r - 1)(c + 2) = 2

                    moduleMatrix(r + 0)(c - 2) = 2
                    moduleMatrix(r + 0)(c - 1) = -2
                    moduleMatrix(r + 0)(c + 0) = 2
                    moduleMatrix(r + 0)(c + 1) = -2
                    moduleMatrix(r + 0)(c + 2) = 2

                    moduleMatrix(r + 1)(c - 2) = 2
                    moduleMatrix(r + 1)(c - 1) = -2
                    moduleMatrix(r + 1)(c + 0) = -2
                    moduleMatrix(r + 1)(c + 1) = -2
                    moduleMatrix(r + 1)(c + 2) = 2

                    moduleMatrix(r + 2)(c - 2) = 2
                    moduleMatrix(r + 2)(c - 1) = 2
                    moduleMatrix(r + 2)(c + 0) = 2
                    moduleMatrix(r + 2)(c + 1) = 2
                    moduleMatrix(r + 2)(c + 2) = 2
                End If
            Next
        Next
    End Sub

End Class


Class AlphanumericEncoder

    Private m_data()
    Private m_charCounter
    Private m_bitCounter

    Public Property Get BitCount()
        BitCount = m_bitCounter
    End Property

    Public Property Get CharCount()
        CharCount = m_charCounter
    End Property

    Public Property Get EncodingMode()
        EncodingMode = ENCODINGMODE_ALPHA_NUMERIC
    End Property

    Public Property Get ModeIndicator()
        ModeIndicator = MODEINDICATOR_ALPAHNUMERIC_VALUE
    End Property

    Public Function Append(ByVal c)
        Dim wd
        wd = ConvertCharCode(c)

        Dim ret

        If m_charCounter Mod 2 = 0 Then
            If m_charCounter = 0 Then
                ReDim m_data(0)
            Else
                ReDim Preserve m_data(UBound(m_data) + 1)
            End If

            m_data(UBound(m_data)) = wd
            ret = 6
        Else
            m_data(UBound(m_data)) = m_data(UBound(m_data)) * 45
            m_data(UBound(m_data)) = m_data(UBound(m_data)) + wd
            ret = 5
        End If

        m_charCounter = m_charCounter + 1
        m_bitCounter = m_bitCounter + ret

        Append = ret
    End Function

    Public Function GetCodewordBitLength(ByVal c)
        If m_charCounter Mod 2 = 0 Then
            GetCodewordBitLength = 6
        Else
            GetCodewordBitLength = 5
        End If
    End Function

    Public Function GetBytes()
        Dim bs
        Set bs = New BitSequence

        Dim bitLength
        bitLength = 11

        Dim i
        For i = 0 To UBound(m_data) - 1
            Call bs.Append(m_data(i), bitLength)
        Next

        If m_charCounter Mod 2 = 0 Then
            bitLength = 11
        Else
            bitLength = 6
        End If

        Call bs.Append(m_data(UBound(m_data)), bitLength)

        GetBytes = bs.GetBytes()
    End Function

    Public Function InSubset(ByVal c)
        Dim ret
        Dim code
        code = Asc(c)

        ' A - Z
        If code >= 65 And code <= 90 Then
            ret = True
        ' 0 - 9
        ElseIf code >= 48 And code <= 57 Then
            ret = True
        ' (Space)
        ElseIf code = 32 Then
            ret = True
        ' $ %
        ElseIf code = 36 Or code = 37 Then
            ret = True
        ' * +
        ElseIf code = 42 Or code = 43 Then
            ret = True
        ' - .
        ElseIf code = 45 Or code = 46 Then
            ret = True
        ' /
        ElseIf code = 47 Then
            ret = True
        ' :
        ElseIf code = 58 Then
            ret = True
        Else
            ret = False
        End If

        InSubset = ret
    End Function

    Public Function InExclusiveSubset(ByVal c)
        Dim ret
        Dim code
        code = Asc(c)

        ' A - Z
        If code >= 65 And code <= 90 Then
            ret = True
        ' (Space)
        ElseIf code = 32 Then
            ret = True
        ' $ %
        ElseIf code = 36 Or code = 37 Then
            ret = True
        ' * +
        ElseIf code = 42 Or code = 43 Then
            ret = True
        ' - .
        ElseIf code = 45 Or code = 46 Then
            ret = True
        ' /
        ElseIf code = 47 Then
            ret = True
        ' :
        ElseIf code = 58 Then
            ret = True
        Else
            ret = False
        End If

        InExclusiveSubset = ret
    End Function

    Public Function ConvertCharCode(ByVal c)
        Dim code
        code = Asc(c)

        ' A - Z
        If code >= 65 And code <= 90 Then
            ConvertCharCode = code - 55
        ' 0 - 9
        ElseIf code >= 48 And code <= 57 Then
            ConvertCharCode = code - 48
        ' (Space)
        ElseIf code = 32 Then
            ConvertCharCode = 36
        ' $ %
        ElseIf code = 36 Or code = 37 Then
            ConvertCharCode = code + 1
        ' * +
        ElseIf code = 42 Or code = 43 Then
            ConvertCharCode = code - 3
        ' - .
        ElseIf code = 45 Or code = 46 Then
            ConvertCharCode = code - 4
        ' /
        ElseIf code = 47 Then
            ConvertCharCode = 43
        ' :
        ElseIf code = 58 Then
            ConvertCharCode = 44
        Else
            ConvertCharCode = -1
        End If
    End Function

End Class


Class BinaryWriter
    Private m_byteTable(255)
    Private m_stream

    Private Sub Class_Initialize()
        Call MakeByteTable

        Set m_stream = CreateObject("ADODB.Stream")
        m_stream.Type = adTypeBinary
        Call m_stream.Open
    End Sub

    Public Property Get Stream()
        Set Stream = m_stream
    End Property

    Public Property Get Size()
        Size = m_stream.Size
    End Property

    Private Sub MakeByteTable()
        Dim sr
        Set sr = CreateObject("ADODB.Stream")
        sr.Type = adTypeText
        sr.Charset = "unicode"
        Call sr.Open

        Dim i
        For i = 0 To 255
            sr.WriteText ChrW(i)
        Next

        sr.Position = 0
        sr.Type = adTypeBinary
        sr.Position = 2

        For i = 0 To 255
            m_byteTable(i) = sr.Read(1)
            sr.Position = sr.Position + 1
        Next

        Call sr.Close
    End Sub

    Public Sub Append(ByVal arg)
        If (VarType(arg) And vbArray) = 0 Then 
            arg = Array(arg)
        End If

        Dim temp

        Dim i
        For i = 0 To Ubound(arg)
            Select Case VarType(arg(i))
                Case vbByte
                    m_stream.Write m_byteTable(arg(i))
                Case vbInteger
                    temp = arg(i) And &HFF&
                    m_stream.Write m_byteTable(temp)

                    temp = (arg(i) And &HFF00&) \ 2 ^ 8
                    m_stream.Write m_byteTable(temp)
                Case vbLong
                    temp = arg(i) And &HFF&
                    m_stream.Write m_byteTable(temp)

                    temp = (arg(i) And &HFF00&) \ 2 ^ 8
                    m_stream.Write m_byteTable(temp)

                    temp = (arg(i) And &HFF0000) \ 2 ^ 16
                    m_stream.Write m_byteTable(temp)

                    temp = (arg(i) And &HFF000000) \ 2 ^ 24
                    m_stream.Write m_byteTable(temp)
                Case Else
                    Call Err.Raise(5)
            End Select
        Next
    End Sub

    Public Sub CopyTo(ByVal destBinaryWriter)
        m_stream.Position = 0
        Call m_stream.CopyTo(destBinaryWriter.Stream)
    End Sub

    Public Sub SaveToFile(ByVal FileName, ByVal SaveOptions)
        Call m_stream.SaveToFile(FileName, SaveOptions)
    End Sub

End Class


Class BitSequence

    Private m_buffer

    Private m_bitCounter
    Private m_space

    Private Sub Class_Initialize()
        Call Clear
    End Sub

    Public Property Get Length()
        Length = m_bitCounter
    End Property

    Public Sub Clear()
        Set m_buffer = CreateObject("Scripting.Dictionary")
        m_bitCounter = 0
        m_space = 0
    End Sub

    Public Sub Append(ByVal data, ByVal bitLength)
        Dim remainingLength
        remainingLength = bitLength

        Dim remainingData
        remainingData = data

        Dim temp

        Do While remainingLength > 0
            If m_space = 0 Then
                m_space = 8
                Call m_buffer.Add(m_buffer.Count, CByte(&H0))
            End If

            temp = m_buffer(m_buffer.Count - 1)

            If m_space < remainingLength Then
                temp = CByte(temp Or remainingData \ (2 ^ (remainingLength - m_space)))

                remainingData = remainingData And ((2 ^ (remainingLength - m_space)) - 1)

                m_bitCounter = m_bitCounter + m_space
                remainingLength = remainingLength - m_space
                m_space = 0
            Else
                temp = CByte(temp Or remainingData * (2 ^ (m_space - remainingLength)))

                m_bitCounter = m_bitCounter + remainingLength
                m_space = m_space - remainingLength
                remainingLength = 0
            End If

            m_buffer(m_buffer.Count - 1) = temp
        Loop
    End Sub

    Public Function GetBytes()
        Dim ret
        ReDim ret(m_buffer.Count - 1)

        Dim i
        For i = 0 To m_buffer.Count - 1
            ret(i) = m_buffer(i)
        Next

        GetBytes = ret
    End Function

End Class


Class ByteEncoder

    Private m_data
    Private m_charCounter
    Private m_bitCounter

    Private m_encAlpha
    Private m_encKanji

    Private Sub Class_Initialize()
        m_data = Empty
        m_charCounter = 0
        m_bitCounter = 0

        Set m_encAlpha = New AlphanumericEncoder
        Set m_encKanji = New KanjiEncoder
    End Sub

    Public Property Get BitCount()
        BitCount = m_bitCounter
    End Property

    Public Property Get CharCount()
        CharCount = m_charCounter
    End Property

    Public Property Get EncodingMode()
        EncodingMode = ENCODINGMODE_EIGHT_BIT_BYTE
    End Property

    Public Property Get ModeIndicator()
        ModeIndicator = MODEINDICATOR_BYTE_VALUE
    End Property

    Public Function Append(ByVal c)
        If m_charCounter = 0 Then
            ReDim m_data(0)
        Else
            ReDim Preserve m_data(UBound(m_data) + 1)
        End If

        Dim wd
        wd = Asc(c) And &HFFFF&

        m_data(UBound(m_data)) = wd

        Dim ret

        If wd > &HFF Then
            m_charCounter = m_charCounter + 2
            ret = 16
        Else
            m_charCounter = m_charCounter + 1
            ret = 8
        End If

        m_bitCounter = m_bitCounter + ret

        Append = ret
    End Function

    Public Function GetCodewordBitLength(ByVal c)
        Dim code
        code = Asc(c) And &HFFFF&

        If code > &HFF Then
            GetCodewordBitLength = 16
        Else
            GetCodewordBitLength = 8
        End If
    End Function

    Public Function GetBytes()
        GetBytes = m_data
    End Function

    Public Function InSubset(ByVal c)
        InSubset = True
    End Function

    Public Function InExclusiveSubset(ByVal c)
        Dim ret

        Dim code
        code = Asc(c) And &HFFFF&

        If code = &H20& Then
            ret = False
        ElseIf code = &H24& Then
            ret = False
        ElseIf code = &H25& Then
            ret = False
        ElseIf code = &H2A& Then
            ret = False
        ElseIf code = &H2B& Then
            ret = False
        ElseIf &H2D& <= code And code <= &H3A& Then
            ret = False
        ElseIf &H41& <= code And code <= &H5A& Then
            ret = False
        ElseIf &H8140& <= code And  code <= &H9FFC& Then
            ret = False
        ElseIf &HE040& <= code And code <= &HEBBF& Then
            ret = False
        Else
            ret = True
        End If

        InExclusiveSubset = ret
    End Function

End Class


Class CharCountIndicator_

    Public Function GetLength(ByVal ver, ByVal encMode)
        If ver >= 1 And ver <= 9 Then
            Select Case encMode
                Case ENCODINGMODE_NUMERIC
                    GetLength = 10
                Case ENCODINGMODE_ALPHA_NUMERIC
                    GetLength = 9
                Case ENCODINGMODE_EIGHT_BIT_BYTE
                    GetLength = 8
                Case ENCODINGMODE_KANJI
                    GetLength = 8
                Case Else
                    Call Err.Raise(5)
            End Select
        ElseIf ver >= 10 And ver <= 26 Then
            Select Case encMode
                Case ENCODINGMODE_NUMERIC
                    GetLength = 12
                Case ENCODINGMODE_ALPHA_NUMERIC
                    GetLength = 11
                Case ENCODINGMODE_EIGHT_BIT_BYTE
                    GetLength = 16
                Case ENCODINGMODE_KANJI
                    GetLength = 10
                Case Else
                    Call Err.Raise(5)
            End Select
        ElseIf ver >= 27 And ver <= 40 Then
            Select Case encMode
                Case ENCODINGMODE_NUMERIC
                    GetLength = 14
                Case ENCODINGMODE_ALPHA_NUMERIC
                    GetLength = 13
                Case ENCODINGMODE_EIGHT_BIT_BYTE
                    GetLength = 16
                Case ENCODINGMODE_KANJI
                    GetLength = 12
                Case Else
                    Call Err.Raise(5)
            End Select
        Else
            Call Err.Raise(5)
        End If
    End Function

End Class


Class Codeword_

    Private m_totalNumbers

    Private Sub Class_Initialize()
        m_totalNumbers = Array( _
            -1, _
            26, 44, 70, 100, 134, 172, 196, 242, 292, 346, _
            404, 466, 532, 581, 655, 733, 815, 901, 991, 1085, _
            1156, 1258, 1364, 1474, 1588, 1706, 1828, 1921, 2051, 2185, _
            2323, 2465, 2611, 2761, 2876, 3034, 3196, 3362, 3532, 3706 _
        )
    End Sub

    Public Function GetTotalNumber(ByVal ver)
        GetTotalNumber = m_totalNumbers(ver)
    End Function

End Class


Class DataCodeword_

    Private m_totalNumbers

    Private Sub Class_Initialize()
        Dim ecLevelL
        ecLevelL = Array( _
               0, _
              19, 34, 55, 80, 108, 136, 156, 194, 232, 274, _
             324, 370, 428, 461, 523, 589, 647, 721, 795, 861, _
             932, 1006, 1094, 1174, 1276, 1370, 1468, 1531, 1631, 1735, _
            1843, 1955, 2071, 2191, 2306, 2434, 2566, 2702, 2812, 2956 _
        )

        Dim ecLevelM
        ecLevelM = Array( _
               0, _
              16, 28, 44, 64, 86, 108, 124, 154, 182, 216, _
             254, 290, 334, 365, 415, 453, 507, 563, 627, 669, _
             714, 782, 860, 914, 1000, 1062, 1128, 1193, 1267, 1373, _
            1455, 1541, 1631, 1725, 1812, 1914, 1992, 2102, 2216, 2334 _
        )

        Dim ecLevelQ
        ecLevelQ = Array( _
               0, _
              13, 22, 34, 48, 62, 76, 88, 110, 132, 154, _
             180, 206, 244, 261, 295, 325, 367, 397, 445, 485, _
             512, 568, 614, 664, 718, 754, 808, 871, 911, 985, _
            1033, 1115, 1171, 1231, 1286, 1354, 1426, 1502, 1582, 1666 _
        )

        Dim ecLevelH
        ecLevelH = Array( _
              0, _
              9, 16, 26, 36, 46, 60, 66, 86, 100, 122, _
            140, 158, 180, 197, 223, 253, 283, 313, 341, 385, _
            406, 442, 464, 514, 538, 596, 628, 661, 701, 745, _
            793, 845, 901, 961, 986, 1054, 1096, 1142, 1222, 1276 _
        )

        m_totalNumbers = Array(ecLevelL, ecLevelM, ecLevelQ, ecLevelH)
    End Sub

    Public Function GetTotalNumber(ByVal ecLevel, ByVal ver)
        GetTotalNumber = m_totalNumbers(ecLevel)(ver)
    End Function

End Class


Class FinderPattern_

    Private m_finderPattern

    Private Sub Class_Initialize()
        m_finderPattern = Array( _
            Array(2, 2, 2, 2, 2, 2, 2), _
            Array(2, -2, -2, -2, -2, -2, 2), _
            Array(2, -2, 2, 2, 2, -2, 2), _
            Array(2, -2, 2, 2, 2, -2, 2), _
            Array(2, -2, 2, 2, 2, -2, 2), _
            Array(2, -2, -2, -2, -2, -2, 2), _
            Array(2, 2, 2, 2, 2, 2, 2) _
        )
    End Sub

    Public Sub Place(ByRef moduleMatrix())
        Dim offset
        offset = (UBound(moduleMatrix) + 1) - (UBound(m_finderPattern) + 1)

        Dim i, j
        Dim v

        For i = 0 To UBound(m_finderPattern)
            For j = 0 To UBound(m_finderPattern(i))
                v = m_finderPattern(i)(j)

                moduleMatrix(i)(j) = v
                moduleMatrix(i)(j + offset) = v
                moduleMatrix(i + offset)(j) = v
            Next
        Next
    End Sub

End Class


Class FormatInfo_

    Private m_formatInfoValues
    Private m_formatInfoMaskArray

    Private Sub Class_Initialize()
        m_formatInfoValues = Array( _
            &H0&, &H537&, &HA6E&, &HF59&, &H11EB&, &H14DC&, &H1B85&, &H1EB2&, &H23D6&, &H26E1&, _
            &H29B8&, &H2C8F&, &H323D&, &H370A&, &H3853&, &H3D64&, &H429B&, &H47AC&, &H48F5&, &H4DC2&, _
            &H5370&, &H5647&, &H591E&, &H5C29&, &H614D&, &H647A&, &H6B23&, &H6E14&, &H70A6&, &H7591&, _
            &H7AC8&, &H7FFF& _
        )

        m_formatInfoMaskArray = Array(0, 1, 0, 0, 1, 0, 0, 0, 0, 0, 1, 0, 1, 0, 1)
    End Sub

    Public Sub Place(ByRef moduleMatrix(), ByVal ecLevel, ByVal maskPatternReference)
        Dim formatInfoValue
        formatInfoValue = GetFormatInfoValue(ecLevel, maskPatternReference)

        Dim temp
        Dim v

        Dim i

        Dim r1
        r1 = 0

        Dim c1
        c1 = UBound(moduleMatrix)

        For i = 0 To 7
            If (formatInfoValue And (2 ^ i)) > 0 Then
                temp = 1 Xor m_formatInfoMaskArray(i)
            Else
                temp = 0 Xor m_formatInfoMaskArray(i)
            End If

            If temp > 0 Then
                v = 3
            Else
                v = -3
            End If

            moduleMatrix(r1)(8) = v
            moduleMatrix(8)(c1) = v

            r1 = r1 + 1
            c1 = c1 - 1

            If r1 = 6 Then
                r1 = r1 + 1
            End If
        Next

        Dim r2
        r2 = UBound(moduleMatrix) - 6

        Dim c2
        c2 = 7

        For i = 8 To 14
            If (formatInfoValue And (2 ^ i)) > 0 Then
                temp = 1 Xor m_formatInfoMaskArray(i)
            Else
                temp = 0 Xor m_formatInfoMaskArray(i)
            End If

            If temp > 0 Then
                v = 3
            Else
                v = -3
            End IF

            moduleMatrix(r2)(8) = v
            moduleMatrix(8)(c2) = v

            r2 = r2 + 1
            c2 = c2 - 1

            If c2 = 6 Then
                c2 = c2 - 1
            End If
        Next
    End Sub

    Public Sub PlaceTempBlank(ByRef moduleMatrix())
        Dim numModulesPerSide
        numModulesPerSide = UBound(moduleMatrix) + 1

        Dim i
        For i = 0 To 8
            If i <> 6 Then
                moduleMatrix(8)(i) = -3
                moduleMatrix(i)(8) = -3
            End If
        Next

        For i = numModulesPerSide - 8 To numModulesPerSide - 1
            moduleMatrix(8)(i) = -3
            moduleMatrix(i)(8) = -3
        Next

        moduleMatrix(numModulesPerSide - 8)(8) = 2
    End Sub

    Private Function GetFormatInfoValue(ByVal ecLevel, ByVal maskPatternReference)
        Dim indicator

        Select Case ecLevel
            Case ERRORCORRECTION_LEVEL_L
                indicator = 1
            Case ERRORCORRECTION_LEVEL_M
                indicator = 0
            Case ERRORCORRECTION_LEVEL_Q
                indicator = 3
            Case ERRORCORRECTION_LEVEL_H
                indicator = 2
            Case Else
                Call Err.Raise(5)
        End Select

        GetFormatInfoValue = m_formatInfoValues((indicator * 2 ^ 3) Or maskPatternReference)
    End Function

End Class


Class GaloisField256_

    Private m_intToExpTable
    Private m_expToIntTable

    Private Sub Class_Initialize()
        m_intToExpTable = Array( _
            -1, 0, 1, 25, 2, 50, 26, 198, 3, 223, 51, 238, 27, 104, 199, 75, 4, 100, 224, 14, 52, 141, _
            239, 129, 28, 193, 105, 248, 200, 8, 76, 113, 5, 138, 101, 47, 225, 36, 15, 33, 53, _
            147, 142, 218, 240, 18, 130, 69, 29, 181, 194, 125, 106, 39, 249, 185, 201, 154, 9, 120, _
            77, 228, 114, 166, 6, 191, 139, 98, 102, 221, 48, 253, 226, 152, 37, 179, 16, 145, 34, 136, _
            54, 208, 148, 206, 143, 150, 219, 189, 241, 210, 19, 92, 131, 56, 70, 64, 30, 66, 182, 163, _
            195, 72, 126, 110, 107, 58, 40, 84, 250, 133, 186, 61, 202, 94, 155, 159, 10, 21, 121, 43, _
            78, 212, 229, 172, 115, 243, 167, 87, 7, 112, 192, 247, 140, 128, 99, 13, 103, 74, 222, 237, _
            49, 197, 254, 24, 227, 165, 153, 119, 38, 184, 180, 124, 17, 68, 146, 217, 35, 32, 137, _
            46, 55, 63, 209, 91, 149, 188, 207, 205, 144, 135, 151, 178, 220, 252, 190, 97, 242, 86, _
            211, 171, 20, 42, 93, 158, 132, 60, 57, 83, 71, 109, 65, 162, 31, 45, 67, 216, 183, 123, 164, _
            118, 196, 23, 73, 236, 127, 12, 111, 246, 108, 161, 59, 82, 41, 157, 85, 170, 251, 96, 134, _
            177, 187, 204, 62, 90, 203, 89, 95, 176, 156, 169, 160, 81, 11, 245, 22, 235, 122, 117, _
            44, 215, 79, 174, 213, 233, 230, 231, 173, 232, 116, 214, 244, 234, 168, 80, 88, 175)

        m_expToIntTable = Array( _
            1, 2, 4, 8, 16, 32, 64, 128, 29, 58, 116, 232, 205, 135, 19, 38, 76, 152, 45, 90, 180, 117, _
            234, 201, 143, 3, 6, 12, 24, 48, 96, 192, 157, 39, 78, 156, 37, 74, 148, 53, 106, 212, 181, _
            119, 238, 193, 159, 35, 70, 140, 5, 10, 20, 40, 80, 160, 93, 186, 105, 210, 185, 111, 222, 161, _
            95, 190, 97, 194, 153, 47, 94, 188, 101, 202, 137, 15, 30, 60, 120, 240, 253, 231, 211, 187, _
            107, 214, 177, 127, 254, 225, 223, 163, 91, 182, 113, 226, 217, 175, 67, 134, 17, 34, 68, 136, _
            13, 26, 52, 104, 208, 189, 103, 206, 129, 31, 62, 124, 248, 237, 199, 147, 59, 118, 236, 197, _
            151, 51, 102, 204, 133, 23, 46, 92, 184, 109, 218, 169, 79, 158, 33, 66, 132, 21, 42, 84, 168, _
            77, 154, 41, 82, 164, 85, 170, 73, 146, 57, 114, 228, 213, 183, 115, 230, 209, 191, 99, 198, _
            145, 63, 126, 252, 229, 215, 179, 123, 246, 241, 255, 227, 219, 171, 75, 150, 49, 98, 196, 149, _
            55, 110, 220, 165, 87, 174, 65, 130, 25, 50, 100, 200, 141, 7, 14, 28, 56, 112, 224, 221, 167, _
            83, 166, 81, 162, 89, 178, 121, 242, 249, 239, 195, 155, 43, 86, 172, 69, 138, 9, 18, 36, 72, 144, _
            61, 122, 244, 245, 247, 243, 251, 235, 203, 139, 11, 22, 44, 88, 176, 125, 250, 233, 207, 131, 27, _
            54, 108, 216, 173, 71, 142, 1)
    End Sub

    Public Function ToExp(ByVal arg)
        ToExp = m_intToExpTable(arg)
    End Function

    Public Function ToInt(ByVal arg)
        ToInt = m_expToIntTable(arg)
    End Function

End Class


Class GeneratorPolynomials_

    Private m_gp(68)

    Private Sub Class_Initialize()
        m_gp(7) = Array(21, 102, 238, 149, 146, 229, 87, 0)
        m_gp(10) = Array(45, 32, 94, 64, 70, 118, 61, 46, 67, 251, 0)
        m_gp(13) = Array(78, 140, 206, 218, 130, 104, 106, 100, 86, 100, 176, 152, 74, 0)
        m_gp(15) = Array(105, 99, 5, 124, 140, 237, 58, 58, 51, 37, 202, 91, 61, 183, 8, 0)
        m_gp(16) = Array(120, 225, 194, 182, 169, 147, 191, 91, 3, 76, 161, 102, 109, 107, 104, 120, 0)
        m_gp(17) = Array(136, 163, 243, 39, 150, 99, 24, 147, 214, 206, 123, 239, 43, 78, 206, 139, 43, 0)
        m_gp(18) = Array(153, 96, 98, 5, 179, 252, 148, 152, 187, 79, 170, 118, 97, 184, 94, 158, 234, 215, 0)
        m_gp(20) = Array(190, 188, 212, 212, 164, 156, 239, 83, 225, 221, 180, 202, 187, 26, 163, 61, 50, 79, 60, 17, 0)
        m_gp(22) = Array(231, 165, 105, 160, 134, 219, 80, 98, 172, 8, 74, 200, 53, 221, 109, 14, 230, 93, 242, 247, 171, 210, 0)
        m_gp(24) = Array(21, 227, 96, 87, 232, 117, 0, 111, 218, 228, 226, 192, 152, 169, 180, 159, 126, 251, 117, 211, 48, 135, 121, 229, 0)
        m_gp(26) = Array(70, 218, 145, 153, 227, 48, 102, 13, 142, 245, 21, 161, 53, 165, 28, 111, 201, 145, 17, 118, 182, 103, 2, 158, 125, 173, 0)
        m_gp(28) = Array(123, 9, 37, 242, 119, 212, 195, 42, 87, 245, 43, 21, 201, 232, 27, 205, 147, 195, 190, 110, 180, 108, 234, 224, 104, 200, 223, 168, 0)
        m_gp(30) = Array(180, 192, 40, 238, 216, 251, 37, 156, 130, 224, 193, 226, 173, 42, 125, 222, 96, 239, 86, 110, 48, 50, 182, 179, 31, 216, 152, 145, 173, 41, 0)
        m_gp(32) = Array(241, 220, 185, 254, 52, 80, 222, 28, 60, 171, 60, 38, 156, 80, 185, 120, 27, 89, 123, 242, 32, 138, 138, 209, 67, 4, 167, 249, 190, 106, 6, 10, 0)
        m_gp(34) = Array(51, 129, 62, 98, 13, 167, 129, 183, 61, 114, 70, 56, 103, 218, 239, 229, 158, 58, 125, 163, 140, 86, 193, 113, 94, 105, 19, 108, 21, 26, 94, 146, 77, 111, 0)
        m_gp(36) = Array(120, 30, 233, 113, 251, 117, 196, 121, 74, 120, 177, 105, 210, 87, 37, 218, 63, 18, 107, 238, 248, 113, 152, 167, 0, 115, 152, 60, 234, 246, 31, 172, 16, 98, 183, 200, 0)
        m_gp(40) = Array(15, 35, 53, 232, 20, 72, 134, 125, 163, 47, 41, 88, 114, 181, 35, 175, 7, 170, 104, 226, 174, 187, 26, 53, 106, 235, 56, 163, 57, 247, 161, 128, 205, 128, 98, 252, 161, 79, 116, 59, 0)
        m_gp(42) = Array(96, 50, 117, 194, 162, 171, 123, 201, 254, 237, 199, 213, 101, 39, 223, 101, 34, 139, 131, 15, 147, 96, 106, 188, 8, 230, 84, 110, 191, 221, 242, 58, 3, 0, 231, 137, 18, 25, 230, 221, 103, 250, 0)
        m_gp(44) = Array(181, 73, 102, 113, 130, 37, 169, 204, 147, 217, 194, 52, 163, 68, 114, 118, 126, 224, 62, 143, 78, 44, 238, 1, 247, 14, 145, 9, 123, 72, 25, 191, 243, 89, 188, 168, 55, 69, 246, 71, 121, 61, 7, 190, 0)
        m_gp(46) = Array(15, 82, 19, 223, 202, 43, 224, 157, 25, 52, 174, 119, 245, 249, 8, 234, 104, 73, 241, 60, 96, 4, 1, 36, 211, 169, 216, 135, 16, 58, 44, 129, 113, 54, 5, 89, 99, 187, 115, 202, 224, 253, 112, 88, 94, 112, 0)
        m_gp(48) = Array(108, 34, 39, 163, 50, 84, 227, 94, 11, 191, 238, 140, 156, 247, 21, 91, 184, 120, 150, 95, 206, 107, 205, 182, 160, 135, 111, 221, 18, 115, 123, 46, 63, 178, 61, 240, 102, 39, 90, 251, 24, 60, 146, 211, 130, 196, 25, 228, 0)
        m_gp(50) = Array(205, 133, 232, 215, 170, 124, 175, 235, 114, 228, 69, 124, 65, 113, 32, 189, 42, 77, 75, 242, 215, 242, 160, 130, 209, 126, 160, 32, 13, 46, 225, 203, 242, 195, 111, 209, 3, 35, 193, 203, 99, 209, 46, 118, 9, 164, 161, 157, 125, 232, 0)
        m_gp(52) = Array(51, 116, 254, 239, 33, 101, 220, 200, 242, 39, 97, 86, 76, 22, 121, 235, 233, 100, 113, 124, 65, 59, 94, 190, 89, 254, 134, 203, 242, 37, 145, 59, 14, 22, 215, 151, 233, 184, 19, 124, 127, 86, 46, 192, 89, 251, 220, 50, 186, 86, 50, 116, 0)
        m_gp(54) = Array(156, 31, 76, 198, 31, 101, 59, 153, 8, 235, 201, 128, 80, 215, 108, 120, 43, 122, 25, 123, 79, 172, 175, 238, 254, 35, 245, 52, 192, 184, 95, 26, 165, 109, 218, 209, 58, 102, 225, 249, 184, 238, 50, 45, 65, 46, 21, 113, 221, 210, 87, 201, 26, 183, 0)
        m_gp(56) = Array(10, 61, 20, 207, 202, 154, 151, 247, 196, 27, 61, 163, 23, 96, 206, 152, 124, 101, 184, 239, 85, 10, 28, 190, 174, 177, 249, 182, 142, 127, 139, 12, 209, 170, 208, 135, 155, 254, 144, 6, 229, 202, 201, 36, 163, 248, 91, 2, 116, 112, 216, 164, 157, 107, 120, 106, 0)
        m_gp(58) = Array(123, 148, 125, 233, 142, 159, 63, 41, 29, 117, 245, 206, 134, 127, 145, 29, 218, 129, 6, 214, 240, 122, 30, 24, 23, 125, 165, 65, 142, 253, 85, 206, 249, 152, 248, 192, 141, 176, 237, 154, 144, 210, 242, 251, 55, 235, 185, 200, 182, 252, 107, 62, 27, 66, 247, 26, 116, 82, 0)
        m_gp(60) = Array(240, 33, 7, 89, 16, 209, 27, 70, 220, 190, 102, 65, 87, 194, 25, 84, 181, 30, 124, 11, 86, 121, 209, 160, 49, 238, 38, 37, 82, 160, 109, 101, 219, 115, 57, 198, 205, 2, 247, 100, 6, 127, 181, 28, 120, 219, 101, 211, 45, 219, 197, 226, 197, 243, 141, 9, 12, 26, 140, 107, 0)
        m_gp(62) = Array(106, 110, 186, 36, 215, 127, 218, 182, 246, 26, 100, 200, 6, 115, 40, 213, 123, 147, 149, 229, 11, 235, 117, 221, 35, 181, 126, 212, 17, 194, 111, 70, 50, 72, 89, 223, 76, 70, 118, 243, 78, 135, 105, 7, 121, 58, 228, 2, 23, 37, 122, 0, 94, 214, 118, 248, 223, 71, 98, 113, 202, 65, 0)
        m_gp(64) = Array(231, 213, 156, 217, 243, 178, 11, 204, 31, 242, 230, 140, 108, 99, 63, 238, 242, 125, 195, 195, 140, 47, 146, 184, 47, 91, 216, 4, 209, 218, 150, 208, 156, 145, 24, 29, 212, 199, 93, 160, 53, 127, 26, 119, 149, 141, 78, 200, 254, 187, 204, 177, 123, 92, 119, 68, 49, 159, 158, 7, 9, 175, 51, 45, 0)
        m_gp(66) = Array(105, 45, 93, 132, 25, 171, 106, 67, 146, 76, 82, 168, 50, 106, 232, 34, 77, 217, 126, 240, 253, 80, 87, 63, 143, 121, 40, 236, 111, 77, 154, 44, 7, 95, 197, 169, 214, 72, 41, 101, 95, 111, 68, 178, 137, 65, 173, 95, 171, 197, 247, 139, 17, 81, 215, 13, 117, 46, 51, 162, 136, 136, 180, 222, 118, 5, 0)
        m_gp(68) = Array(238, 163, 8, 5, 3, 127, 184, 101, 27, 235, 238, 43, 198, 175, 215, 82, 32, 54, 2, 118, 225, 166, 241, 137, 125, 41, 177, 52, 231, 95, 97, 199, 52, 227, 89, 160, 173, 253, 84, 15, 84, 93, 151, 203, 220, 165, 202, 60, 52, 133, 205, 190, 101, 84, 150, 43, 254, 32, 160, 90, 70, 77, 93, 224, 33, 223, 159, 247, 0)
    End Sub

    Public Function Item(ByVal numECCodewords)
        If IsEmpty(m_gp(numECCodewords)) Then Call Err.Raise(5)

        Item = m_gp(numECCodewords)
    End Function

End Class


Class KanjiEncoder

    Private m_data()
    Private m_charCounter
    Private m_bitCounter

    Public Property Get BitCount()
        BitCount = m_bitCounter
    End Property

    Public Property Get CharCount()
        CharCount = m_charCounter
    End Property

    Public Property Get EncodingMode()
        EncodingMode = ENCODINGMODE_KANJI
    End Property

    Public Property Get ModeIndicator()
        ModeIndicator = MODEINDICATOR_KANJI_VALUE
    End Property

    Public Function Append(ByVal c)
        Dim wd
        wd = Asc(c) And &HFFFF&

        If &H8140& <= wd And wd <= &H9FFC& Then
            wd = wd - &H8140&
        ElseIf &HE040& <= wd And wd <= &HEBBF& Then
            wd = wd - &HC140&
        Else
            Call Err.Raise(5)
        End If

        wd = ((wd \ 2 ^ 8) * &HC0&) + (wd And &HFF&)
        If m_charCounter = 0 Then
            ReDim m_data(0)
        Else
            ReDim Preserve m_data(UBound(m_data) + 1)
        End If

        m_data(UBound(m_data)) = wd

        m_charCounter = m_charCounter + 1
        m_bitCounter = m_bitCounter + 13

        Append = 13
    End Function

    Public Function GetCodewordBitLength(ByVal c)
        GetCodewordBitLength = 13
    End Function

    Public Function GetBytes()
        Dim bs
        Set bs = New BitSequence

        Dim i
        For i = 0 To UBound(m_data)
            Call bs.Append(m_data(i), 13)
        Next

        GetBytes = bs.GetBytes()
    End Function

    Public Function InSubset(ByVal c)
        Dim code
        code = Asc(c) And &HFFFF&

        Dim lsb
        lsb = code And &HFF&

        If &H8140& <= code And code <= &H9FFC& Or _
           &HE040& <= code And code <= &HEBBF& Then
            InSubset = &H40&  <= lsb And lsb <= &HFC& And _
                       lsb <> &H7F&
        Else
            InSubset = False
        End If
    End Function

    Public Function InExclusiveSubset(ByVal c)
        InExclusiveSubset = InSubset(c)
    End Function

End Class


Class List

    Private m_items

    Private Sub Class_Initialize()
        m_items = Array()
    End Sub

    Public Sub Add(arg)
        ReDim Preserve m_items(UBound(m_items) + 1)

        If VarType(arg) = vbObject Then
            Set m_items(UBound(m_items)) = arg
        Else
            m_items(UBound(m_items)) = arg
        End If
    End Sub

    Public Property Get Count()
        Count = UBound(m_items) + 1
    End Property

    Public Property Get Item(ByVal idx)
        If VarType(m_items(idx)) = vbObject Then
            Set Item = m_items(idx)
        Else
            Item = m_items(idx)
        End If
    End Property

    Public Property Get Items()
        Items = m_items
    End Property

End Class


Class Masking_

    Public Function Apply(ByRef moduleMatrix(), ByVal ver, ByVal ecLevel)
        Dim maskPatternReference
        maskPatternReference = SelectMaskPattern(moduleMatrix, ver, ecLevel)

        Call Mask(moduleMatrix, maskPatternReference)

        Apply = maskPatternReference
    End Function

    Private Function SelectMaskPattern(ByRef moduleMatrix(), ByVal ver, ByVal ecLevel)
        Dim minPenalty
        minPenalty = &H7FFFFFFF

        Dim ret
        ret = 0

        Dim temp
        Dim penalty
        Dim maskPatternReference

        For maskPatternReference = 0 To 7
            temp = moduleMatrix
            Call Mask(temp, maskPatternReference)
            Call FormatInfo.Place(temp, ecLevel, maskPatternReference)

            If ver >= 7 Then
                Call VersionInfo.Place(temp, ver)
            End If

            penalty = MaskingPenaltyScore.CalcTotal(temp)

            If penalty < minPenalty Then
                minPenalty = penalty
                ret = maskPatternReference
            End If
        Next

        SelectMaskPattern = ret
    End Function

    Private Sub Mask(ByRef moduleMatrix(), ByVal maskPatternReference)
        Dim condition
        Set condition = GetCondition(maskPatternReference)

        Dim r, c

        For r = 0 To UBound(moduleMatrix)
            For c = 0 To UBound(moduleMatrix(r))
                If Abs(moduleMatrix(r)(c)) = 1 Then
                    If condition.Evaluate(r, c) Then
                        moduleMatrix(r)(c) = moduleMatrix(r)(c) * -1
                    End If
                End If
            Next
        Next
    End Sub

    Private Function GetCondition(ByVal maskPatternReference)
        Dim ret

        Select Case maskPatternReference
            Case 0
                Set ret = New MaskingCondition0
            Case 1
                Set ret = New MaskingCondition1
            Case 2
                Set ret = New MaskingCondition2
            Case 3
                Set ret = New MaskingCondition3
            Case 4
                Set ret = New MaskingCondition4
            Case 5
                Set ret = New MaskingCondition5
            Case 6
                Set ret = New MaskingCondition6
            Case 7
                Set ret = New MaskingCondition7
            Case Else
                Call Err.Raise(5)
        End Select

        Set GetCondition = ret
    End Function

End Class


Class MaskingCondition0

    Public Function Evaluate(ByVal r, ByVal c)
        Evaluate = (r + c) Mod 2 = 0
    End Function

End Class


Class MaskingCondition1

    Public Function Evaluate(ByVal r, ByVal c)
        Evaluate = r Mod 2 = 0
    End Function

End Class


Class MaskingCondition2

    Public Function Evaluate(ByVal r, ByVal c)
        Evaluate = c Mod 3 = 0
    End Function

End Class


Class MaskingCondition3

    Public Function Evaluate(ByVal r, ByVal c)
        Evaluate = (r + c) Mod 3 = 0
    End Function

End Class


Class MaskingCondition4

    Public Function Evaluate(ByVal r, ByVal c)
        Evaluate = ((r \ 2) + (c \ 3)) Mod 2 = 0
    End Function

End Class


Class MaskingCondition5

    Public Function Evaluate(ByVal r, ByVal c)
        Evaluate = ((r * c) Mod 2 + (r * c) Mod 3) = 0
    End Function

End Class


Class MaskingCondition6

    Public Function Evaluate(ByVal r, ByVal c)
        Evaluate = ((r * c) Mod 2 + (r * c) Mod 3) Mod 2 = 0
    End Function

End Class


Class MaskingCondition7

    Public Function Evaluate(ByVal r, ByVal c)
        Evaluate = ((r + c) Mod 2 + (r * c) Mod 3) Mod 2 = 0
    End Function

End Class


Class MaskingPenaltyScore_

    Public Function CalcTotal(ByRef moduleMatrix())
        Dim total
        Dim penalty

        penalty = CalcAdjacentModulesInSameColor(moduleMatrix)
        total = total + penalty

        penalty = CalcBlockOfModulesInSameColor(moduleMatrix)
        total = total + penalty

''        penalty = CalcModuleRatio(moduleMatrix)
''        total = total + penalty

        penalty = CalcProportionOfDarkModules(moduleMatrix)
        total = total + penalty

        CalcTotal = total
    End Function

    Private Function CalcAdjacentModulesInSameColor(ByRef moduleMatrix())
        Dim penalty
        penalty = 0

        penalty = penalty + CalcAdjacentModulesInRowInSameColor(moduleMatrix)
        penalty = penalty + CalcAdjacentModulesInRowInSameColor(MatrixRotate90(moduleMatrix))

        CalcAdjacentModulesInSameColor = penalty
    End Function

    Private Function CalcAdjacentModulesInRowInSameColor(ByRef moduleMatrix())
        Dim penalty
        penalty = 0

        Dim r, c
        Dim cnt

        For r = 0 To UBound(moduleMatrix)
            cnt = 1

            For c = 0 To UBound(moduleMatrix(r)) - 1
                If (moduleMatrix(r)(c) > 0) = (moduleMatrix(r)(c + 1) > 0) Then
                    cnt = cnt + 1
                Else
                    If cnt >= 5 Then
                        penalty = penalty + (3 + (cnt - 5))
                    End If

                    cnt = 1
                End If
            Next

            If cnt >= 5 Then
                penalty = penalty + (3 + (cnt - 5))
            End If
        Next

        CalcAdjacentModulesInRowInSameColor = penalty
    End Function

    Private Function CalcBlockOfModulesInSameColor(ByRef moduleMatrix())
        Dim penalty
        Dim isSameColor
        Dim r, c
        Dim temp

        For r = 0 To UBound(moduleMatrix) - 1
            For c = 0 To UBound(moduleMatrix(r)) - 1
                temp = moduleMatrix(r)(c) > 0
                isSameColor = True

                isSameColor = isSameColor And (moduleMatrix(r + 0)(c + 1) > 0 = temp)
                isSameColor = isSameColor And (moduleMatrix(r + 1)(c + 0) > 0 = temp)
                isSameColor = isSameColor And (moduleMatrix(r + 1)(c + 1) > 0 = temp)

                If isSameColor Then
                    penalty = penalty + 3
                End If
            Next
        Next

        CalcBlockOfModulesInSameColor = penalty
    End Function

    Private Function CalcModuleRatio(ByRef moduleMatrix())
        Dim moduleMatrixTemp
        moduleMatrixTemp = QuietZone.Place(moduleMatrix)

        Dim penalty
        penalty = 0

        penalty = penalty + CalcModuleRatioInRow(moduleMatrixTemp)
        penalty = penalty + CalcModuleRatioInRow(MatrixRotate90(moduleMatrixTemp))

        CalcModuleRatio = penalty
    End Function

    Private Function CalcModuleRatioInRow(ByRef moduleMatrix())
        Dim penalty
        penalty = 0

        Dim r, c
        Dim cols
        Dim maxIdx
        Dim startIndexes

        Dim i
        Dim idx
        Dim ratio

        For r = 0 To UBound(moduleMatrix)
            cols = moduleMatrix(r)
            maxIdx = Ubound(cols)
            Set startIndexes = New List

            Call startIndexes.Add(0)

            For c = 0 To maxIdx - 2
                If cols(c) > 0 And cols(c + 1) <= 0 Then
                    Call startIndexes.Add(c + 1)
                End If
            Next

            For i = 0 To startIndexes.Count - 1
                idx = startIndexes.Item(i)
                Set ratio = New ModuleRatio

                Do While idx <= maxIdx
                    If cols(idx) > 0 Then Exit Do
                    ratio.PreLightRatio4 = ratio.PreLightRatio4 + 1
                    idx = idx + 1                    
                Loop

                Do While idx <= maxIdx
                    If cols(idx) <= 0 Then Exit Do
                    ratio.PreDarkRatio1 = ratio.PreDarkRatio1 + 1
                    idx = idx + 1
                Loop

                Do While idx <= maxIdx
                    If cols(idx) > 0 Then Exit Do
                    ratio.PreLightRatio1 = ratio.PreLightRatio1 + 1
                    idx = idx + 1                   
                Loop

                If ratio.PreDarkRatio1 = ratio.PreLightRatio1 Then
                    Do While idx <= maxIdx
                        If cols(idx) <= 0 Then Exit Do
                        ratio.CenterDarkRatio3 = ratio.CenterDarkRatio3 + 1
                        idx = idx + 1                       
                    Loop

                    If (ratio.PreLightRatio1 * 3) = ratio.CenterDarkRatio3 Then
                        Do While idx <= maxIdx
                            If cols(idx) > 0 Then Exit Do
                            ratio.FolLightRatio1 = ratio.FolLightRatio1 + 1
                            idx = idx + 1                            
                        Loop

                        If ratio.CenterDarkRatio3 = (ratio.FolLightRatio1 * 3) Then 
                            Do While idx <= maxIdx
                                If cols(idx) <= 0 Then Exit Do
                                ratio.FolDarkRatio1 = ratio.FolDarkRatio1 + 1
                                idx = idx + 1                                
                            Loop

                            If ratio.FolLightRatio1 = ratio.FolDarkRatio1 Then
                                Do While idx <= maxIdx
                                    If cols(idx) > 0 Then Exit Do
                                    ratio.FolLightRatio4 = ratio.FolLightRatio4 + 1
                                    idx = idx + 1                                    
                                Loop
                            End If
                        End If
                    End If
                End If

                If ratio.PenaltyImposed() Then
                    penalty = penalty + 40
                End If
            Next
        Next

        CalcModuleRatioInRow = penalty
    End Function

    Private Function CalcProportionOfDarkModules(ByRef moduleMatrix())
        Dim darkCount
        Dim r, c

        For r = 0 To UBound(moduleMatrix)
            For c = 0 To UBound(moduleMatrix(r))
                If moduleMatrix(r)(c) > 0 Then
                    darkCount = darkCount + 1
                End If
            Next
        Next

        Dim numModules
        numModules = (UBound(moduleMatrix) + 1) ^ 2

        Dim temp
        temp = Int((darkCount / numModules * 100) + 1)
        temp = Abs(temp - 50)
        temp = (temp + 4) \ 5

        CalcProportionOfDarkModules = temp * 10
    End Function

    Private Function MatrixRotate90(ByRef arg())
        Dim ret()
        ReDim ret(UBound(arg(0)))

        Dim i, j
        Dim cols()

        For i = 0 To UBound(ret)
            ReDim cols(UBound(arg))
            ret(i) = cols
        Next

        Dim k
        k = UBound(ret)

        For i = 0 To UBound(ret)
            For j = 0 To UBound(ret(i))
                ret(i)(j) = arg(j)(k - i)
            Next
        Next

        MatrixRotate90 = ret
    End Function

End Class


Class Module_

    Public Function GetNumModulesPerSide(ByVal ver)
        GetNumModulesPerSide = 17 + ver * 4
    End Function

End Class


Class ModuleRatio

    Public PreLightRatio4
    Public PreDarkRatio1
    Public PreLightRatio1
    Public CenterDarkRatio3
    Public FolLightRatio1
    Public FolDarkRatio1
    Public FolLightRatio4

    Public Function PenaltyImposed()
        If PreDarkRatio1 = 0 Then
            PenaltyImposed = False
            Exit Function
        End If

        If PreDarkRatio1 = PreLightRatio1 And _
           PreDarkRatio1 = FolLightRatio1 And _
           PreDarkRatio1 = FolDarkRatio1 And _
           PreDarkRatio1 * 3 = CenterDarkRatio3 Then

            PenaltyImposed = PreLightRatio4 >= PreDarkRatio1 * 4 Or _
                             FolLightRatio4 >= PreDarkRatio1 * 4
        Else
            PenaltyImposed = False
        End If
    End Function

End Class


Class NumericEncoder

    Private m_data()
    Private m_charCounter
    Private m_bitCounter

    Public Property Get BitCount()
        BitCount = m_bitCounter
    End Property

    Public Property Get CharCount()
        CharCount = m_charCounter
    End Property

    Public Property Get EncodingMode()
        EncodingMode = ENCODINGMODE_NUMERIC
    End Property

    Public Property Get ModeIndicator()
        ModeIndicator = MODEINDICATOR_NUMERIC_VALUE
    End Property

    Public Function Append(ByVal c)
        Dim ret

        If m_charCounter Mod 3 = 0 Then
            If m_charCounter = 0 Then
                ReDim m_data(0)
            Else
                ReDim Preserve m_data(UBound(m_data) + 1)
            End If

            m_data(UBound(m_data)) = CLng(c)
            ret = 4
        Else
            m_data(UBound(m_data)) = m_data(UBound(m_data)) * 10 + CLng(c)
            ret = 3
        End If

        m_charCounter = m_charCounter + 1
        m_bitCounter = m_bitCounter + ret

        Append = ret
    End Function

    Public Function GetCodewordBitLength(ByVal c)
        If m_charCounter Mod 3 = 0 Then
            GetCodewordBitLength = 4
        Else
            GetCodewordBitLength = 3
        End If
    End Function

    Public Function GetBytes()
        Dim bs
        Set bs = New BitSequence

        Dim i
        For i = 0 To UBound(m_data) - 1
            Call bs.Append(m_data(i), 10)
        Next

        Select Case m_charCounter Mod 3
            Case 1
                Call bs.Append(m_data(UBound(m_data)), 4)
            Case 2
                Call bs.Append(m_data(UBound(m_data)), 7)
            Case Else
                Call bs.Append(m_data(UBound(m_data)), 10)
        End Select

        GetBytes = bs.GetBytes()
    End Function

    Public Function InSubset(ByVal c)
        InSubset = "0" <= c And c <= "9"
    End Function

    Public Function InExclusiveSubset(ByVal c)
        InExclusiveSubset = InSubset(c)
    End Function

End Class


Class QuietZone_

    Public Function Place(ByRef moduleMatrix())
        Const QUIET_ZONE_WIDTH = 4

        Dim ret()
        ReDim ret(UBound(moduleMatrix) + QUIET_ZONE_WIDTH * 2)

        Dim i
        Dim cols()

        For i = 0 To UBound(ret)
            ReDim cols(UBound(ret))
            ret(i) = cols
        Next

        Dim r
        Dim c

        For r = 0 To UBound(moduleMatrix)
            For c = 0 To UBound(moduleMatrix(r))
                ret(r + QUIET_ZONE_WIDTH)(c + QUIET_ZONE_WIDTH) = moduleMatrix(r)(c)
            Next
        Next

        Place = ret
    End Function

End Class


Class RemainderBit_

    Public Sub Place(ByRef moduleMatrix())
        Dim r, c

        For r = 0 To UBound(moduleMatrix)
            For c = 0 To UBound(moduleMatrix(r))
                If moduleMatrix(r)(c) = 0 Then
                    moduleMatrix(r)(c) = -1
                End If
            Next
        Next
    End Sub

End Class


Class RSBlock_

    Private m_totalNumbers

    Private Sub Class_Initialize()
        m_totalNumbers = Array( _
            Array(0, _
                  1, 1, 1, 1, 1, 2, 2, 2, 2, 4, _
                  4, 4, 4, 4, 6, 6, 6, 6, 7, 8, _
                  8, 9, 9, 10, 12, 12, 12, 13, 14, 15, _
                  16, 17, 18, 19, 19, 20, 21, 22, 24, 25), _
            Array(0, _
                  1, 1, 1, 2, 2, 4, 4, 4, 5, 5, _
                  5, 8, 9, 9, 10, 10, 11, 13, 14, 16, _
                  17, 17, 18, 20, 21, 23, 25, 26, 28, 29, _
                  31, 33, 35, 37, 38, 40, 43, 45, 47, 49), _
            Array(0, _
                  1, 1, 2, 2, 4, 4, 6, 6, 8, 8, _
                  8, 10, 12, 16, 12, 17, 16, 18, 21, 20, _
                  23, 23, 25, 27, 29, 34, 34, 35, 38, 40, _
                  43, 45, 48, 51, 53, 56, 59, 62, 65, 68), _
            Array(0, _
                  1, 1, 2, 4, 4, 4, 5, 6, 8, 8, _
                  11, 11, 16, 16, 18, 16, 19, 21, 25, 25, _
                  25, 34, 30, 32, 35, 37, 40, 42, 45, 48, _
                  51, 54, 57, 60, 63, 66, 70, 74, 77, 81) _
        )
    End Sub

    Public Function GetTotalNumber(ByVal ecLevel, ByVal ver, ByVal preceding)
        Dim dataWordCapacity
        Dim blockCount

        dataWordCapacity = DataCodeword.GetTotalNumber(ecLevel, ver)
        blockCount = m_totalNumbers(ecLevel)

        If preceding Then
            GetTotalNumber = blockCount(ver) - (dataWordCapacity Mod blockCount(ver))
        Else
            GetTotalNumber = dataWordCapacity Mod blockCount(ver)
        End If
    End Function

    Public Function GetNumberDataCodewords(ByVal ecLevel, ByVal ver, ByVal preceding)
        Dim ret

        Dim numDataCodewords
        numDataCodewords = DataCodeword.GetTotalNumber(ecLevel, ver)

        Dim numBlocks
        numBlocks = m_totalNumbers(ecLevel)(ver)

        Dim numPreBlockCodewords
        numPreBlockCodewords = numDataCodewords \ numBlocks

        Dim numPreBlocks
        Dim numFolBlocks

        If preceding Then
            ret = numPreBlockCodewords
        Else
            numPreBlocks = GetTotalNumber(ecLevel, ver, True)
            numFolBlocks = GetTotalNumber(ecLevel, ver, False)

            If numFolBlocks > 0 Then
                ret = (numDataCodewords - numPreBlockCodewords * numPreBlocks) \ numFolBlocks
            Else
                ret = 0
            End If
        End If

        GetNumberDataCodewords = ret
    End Function

    Public Function GetNumberECCodewords(ByVal ecLevel, ByVal ver)
        Dim numDataCodewords
        numDataCodewords = DataCodeword.GetTotalNumber(ecLevel, ver)

        Dim numBlocks
        numBlocks = m_totalNumbers(ecLevel)(ver)

        GetNumberECCodewords = _
            (Codeword.GetTotalNumber(ver) \ numBlocks) - _
                (numDataCodewords \ numBlocks)
    End Function

End Class


Class Separator_

    Public Sub Place(ByRef moduleMatrix())
        Dim offset
        offset = UBound(moduleMatrix) - 7

        Dim i
        For i = 0 To 7
             moduleMatrix(i)(7) = -2
             moduleMatrix(7)(i) = -2

             moduleMatrix(offset + i)(7) = -2
             moduleMatrix(offset + 0)(i) = -2

             moduleMatrix(i)(offset + 0) = -2
             moduleMatrix(7)(offset + i) = -2
         Next
    End Sub

End Class


Class Symbol

    Private m_parent

    Private m_position

    Private m_currEncoder
    Private m_currEncodingMode
    Private m_currVersion

    Private m_dataBitCapacity
    Private m_dataBitCounter

    Private m_segments
    Private m_segmentCounter

    Private Sub Class_Initialize()
        Set m_segments = New List
        Set m_segmentCounter = CreateObject("Scripting.Dictionary")
    End Sub

    Public Sub Init(ByVal parentObj)
        Set m_parent = parentObj

        m_position = parentObj.Count

        Set m_currEncoder = Nothing
        m_currEncodingMode = ENCODINGMODE_UNKNOWN
        m_currVersion = parentObj.MinVersion

        m_dataBitCapacity = 8 * DataCodeword.GetTotalNumber( _
            parentObj.ErrorCorrectionLevel, parentObj.MinVersion)

        m_dataBitCounter = 0

        Call m_segmentCounter.Add(ENCODINGMODE_NUMERIC, 0)
        Call m_segmentCounter.Add(ENCODINGMODE_ALPHA_NUMERIC, 0)
        Call m_segmentCounter.Add(ENCODINGMODE_EIGHT_BIT_BYTE, 0)
        Call m_segmentCounter.Add(ENCODINGMODE_KANJI, 0)
        
        If parentObj.StructuredAppendAllowed Then
            m_dataBitCapacity = m_dataBitCapacity - STRUCTUREDAPPEND_HEADER_LENGTH
        End If
    End Sub

    Public Property Get Parent()
        Set Parent = m_parent
    End Property

    Public Property Get Version()
        Version = m_currVersion
    End Property

    Public Property Get CurrentEncodingMode()
        CurrentEncodingMode = m_currEncodingMode
    End Property

    Public Function TryAppend(ByVal c)
        Dim bitLength
        bitLength = m_currEncoder.GetCodewordBitLength(c)

        Do While (m_dataBitCapacity < m_dataBitCounter + bitLength)
            If m_currVersion >= m_parent.MaxVersion Then
                TryAppend = False
                Exit Function
            End If

            Call SelectVersion
        Loop

        Call m_currEncoder.Append(c)
        m_dataBitCounter = m_dataBitCounter + bitLength
        Call m_parent.UpdateParity(c)

        TryAppend = True
    End Function

    Public Function TrySetEncodingMode(ByVal encMode, ByVal c)
        Dim encoder
        Set encoder = CreateEncoder(encMode)

        Dim bitLength
        bitLength = encoder.GetCodewordBitLength(c)

        Do While (m_dataBitCapacity < _
                    m_dataBitCounter + _
                    MODEINDICATOR_LENGTH + _
                    CharCountIndicator.GetLength(m_currVersion, encMode) + _
                    bitLength)

            If m_currVersion >= m_parent.MaxVersion Then
                TrySetEncodingMode = False
                Exit Function
            End If

            Call SelectVersion
        Loop

        m_dataBitCounter = m_dataBitCounter + _
                           MODEINDICATOR_LENGTH + _
                           CharCountIndicator.GetLength(m_currVersion, encMode)

        Set m_currEncoder = encoder
        Call m_segments.Add(encoder)
        m_segmentCounter(encMode) = m_segmentCounter(encMode) + 1
        m_currEncodingMode = encMode

        TrySetEncodingMode = True
    End Function

    Private Sub SelectVersion()
        Dim encMode

        For Each encMode In m_segmentCounter.Keys
            Dim num
            num = m_segmentCounter(encMode)

            m_dataBitCounter = m_dataBitCounter + _
                               num * CharCountIndicator.GetLength( _
                                    m_currVersion + 1, encMode) - _
                               num * CharCountIndicator.GetLength( _
                                    m_currVersion + 0, encMode)
        Next

        m_currVersion = m_currVersion + 1
        m_dataBitCapacity = 8 * DataCodeword.GetTotalNumber( _
            m_parent.ErrorCorrectionLevel, m_currVersion)
        m_parent.MinVersion = m_currVersion

        If m_parent.StructuredAppendAllowed Then
            m_dataBitCapacity = m_dataBitCapacity - STRUCTUREDAPPEND_HEADER_LENGTH
        End If
    End Sub

    Private Function BuildDataBlock()
        Dim dataBytes
        dataBytes = GetMessageBytes()

        Dim numPreBlocks
        numPreBlocks = RSBlock.GetTotalNumber( _
            m_parent.ErrorCorrectionLevel, m_currVersion, True)

        Dim numFolBlocks
        numFolBlocks = RSBlock.GetTotalNumber( _
            m_parent.ErrorCorrectionLevel, m_currVersion, False)

        Dim ret()
        ReDim ret(numPreBlocks + numFolBlocks - 1)

        Dim dataIdx
        dataIdx = 0

        Dim numPreBlockDataCodewords
        numPreBlockDataCodewords = RSBlock.GetNumberDataCodewords( _
            m_parent.ErrorCorrectionLevel, m_currVersion, True)

        Dim data()
        Dim i, j

        For i = 0 To numPreBlocks - 1
            ReDim data(numPreBlockDataCodewords - 1)

            For j = 0 To UBound(data)
                data(j) = dataBytes(dataIdx)
                dataIdx = dataIdx + 1
            Next

            ret(i) = data
        Next

        Dim numFolBlockDataCodewords
        numFolBlockDataCodewords = RSBlock.GetNumberDataCodewords( _
            m_parent.ErrorCorrectionLevel, m_currVersion, False)

        For i = numPreBlocks To numPreBlocks + numFolBlocks - 1
            ReDim data(numFolBlockDataCodewords - 1)

            For j = 0 To UBound(data)
                data(j) = dataBytes(dataIdx)
                dataIdx = dataIdx + 1
            Next

            ret(i) = data
        Next

        BuildDataBlock = ret
    End Function

    Private Function BuildErrorCorrectionBlock(ByRef dataBlock())
        Dim i, j

        Dim numECCodewords
        numECCodewords = RSBlock.GetNumberECCodewords( _
            m_parent.ErrorCorrectionLevel, m_currVersion)

        Dim numPreBlocks
        numPreBlocks = RSBlock.GetTotalNumber( _
            m_parent.ErrorCorrectionLevel, m_currVersion, True)

        Dim numFolBlocks
        numFolBlocks = RSBlock.GetTotalNumber( _
            m_parent.ErrorCorrectionLevel, m_currVersion, False)

        Dim ret()
        ReDim ret(numPreBlocks + numFolBlocks - 1)

        Dim eccDataTmp()
        ReDim eccDataTmp(numECCodewords - 1)

        For i = 0 To UBound(ret)
            ret(i) = eccDataTmp
        Next

        Dim gp
        gp = GeneratorPolynomials.Item(numECCodewords)

        Dim eccIdx
        Dim blockIdx
        Dim data()
        Dim exp

        For blockIdx = 0 To UBound(dataBlock)
            ReDim data(UBound(dataBlock(blockIdx)) + UBound(ret(blockIdx)) + 1)
            eccIdx = UBound(data)

            For i = 0 To UBound(dataBlock(blockIdx))
                data(eccIdx) = dataBlock(blockIdx)(i)
                eccIdx = eccIdx - 1
            Next

            For i = UBound(data) To numECCodewords Step -1
                If data(i) > 0 Then
                    exp = GaloisField256.ToExp(data(i))
                    eccIdx = i

                    For j = UBound(gp) To 0 Step -1
                        data(eccIdx) = data(eccIdx) Xor _
                                       GaloisField256.ToInt((gp(j) + exp) Mod 255)
                        eccIdx = eccIdx - 1
                    Next
                End If
            Next

            eccIdx = numECCodewords - 1

            For i = 0 To UBound(ret(blockIdx))
                ret(blockIdx)(i) = data(eccIdx)
                eccIdx = eccIdx - 1
            Next
        Next

        BuildErrorCorrectionBlock = ret
    End Function

    Private Function GetEncodingRegionBytes()
        Dim dataBlock
        dataBlock = BuildDataBlock()

        Dim ecBlock
        ecBlock = BuildErrorCorrectionBlock(dataBlock)

        Dim numCodewords
        numCodewords = Codeword.GetTotalNumber(m_currVersion)

        Dim numDataCodewords
        numDataCodewords = DataCodeword.GetTotalNumber( _
            m_parent.ErrorCorrectionLevel, m_currVersion)

        Dim ret()
        ReDim ret(numCodewords - 1)

        Dim r, c

        Dim idx
        idx = 0

        Dim n
        n = 0

        Do While idx < numDataCodewords
            r = n Mod (UBound(dataBlock) + 1)
            c = n \ (UBound(dataBlock) + 1)

            If c <= UBound(dataBlock(r)) Then
                ret(idx) = dataBlock(r)(c)
                idx = idx + 1
            End If

            n = n + 1
        Loop

        n = 0

        Do While idx < numCodewords
            r = n Mod (UBound(ecBlock) + 1)
            c = n \ (UBound(ecBlock) + 1)

            If c <= UBound(ecBlock(r)) Then
                ret(idx) = ecBlock(r)(c)
                idx = idx + 1
            End If

            n = n + 1
        Loop

        GetEncodingRegionBytes = ret
    End Function

    Private Function GetMessageBytes()
        Dim bs
        Set bs = New BitSequence

        If m_parent.Count > 1 Then
            Call WriteStructuredAppendHeader(bs)
        End If

        Call WriteSegments(bs)
        Call WriteTerminator(bs)
        Call WritePaddingBits(bs)
        Call WritePadCodewords(bs)

        GetMessageBytes = bs.GetBytes()
    End Function

    Private Sub WriteStructuredAppendHeader(ByVal bs)
        Call bs.Append(MODEINDICATOR_STRUCTURED_APPEND_VALUE, _
                       MODEINDICATOR_LENGTH)
        Call bs.Append(m_position, _
                       SYMBOLSEQUENCEINDICATOR_POSITION_LENGTH)
        Call bs.Append(m_parent.Count - 1, _
                       SYMBOLSEQUENCEINDICATOR_TOTAL_NUMBER_LENGTH)
        Call bs.Append(m_parent.StructuredAppendParity, _
                       STRUCTUREDAPPEND_PARITY_DATA_LENGTH)
    End Sub

    Private Sub WriteSegments(ByVal bs)
        Dim i
        Dim data
        Dim codewordBitLength

        Dim segment

        For Each segment In m_segments.Items()
            Call bs.Append(segment.ModeIndicator, MODEINDICATOR_LENGTH)
            Call bs.Append(segment.CharCount, _
                           CharCountIndicator.GetLength( _
                                m_currVersion, segment.EncodingMode))

            data = segment.GetBytes()

            For i = 0 To UBound(data) - 1
                Call bs.Append(data(i), 8)
            Next

            codewordBitLength = segment.BitCount Mod 8

            If codewordBitLength = 0 Then
                codewordBitLength = 8
            End If

            Call bs.Append(data(UBound(data)) \ _
                           2 ^ (8 - codewordBitLength), codewordBitLength)
        Next
    End Sub

    Private Sub WriteTerminator(ByVal bs)
        Dim terminatorLength
        terminatorLength = m_dataBitCapacity - m_dataBitCounter

        If terminatorLength > MODEINDICATOR_LENGTH Then
            terminatorLength = MODEINDICATOR_LENGTH
        End If

        Call bs.Append(MODEINDICATOR_TERMINATOR_VALUE, terminatorLength)
    End Sub

    Private Sub WritePaddingBits(ByVal bs)

        If bs.Length Mod 8 > 0 Then
            Call bs.Append(&H0, 8 - (bs.Length Mod 8))
        End If
    End Sub

    Private Sub WritePadCodewords(ByVal bs)
        Dim numDataCodewords
        numDataCodewords = DataCodeword.GetTotalNumber( _
            m_parent.ErrorCorrectionLevel, m_currVersion)

        Dim flag
        flag = True

        Dim v

        Do While bs.Length < 8 * numDataCodewords
            If flag Then
                v = 236
            Else
                v = 17
            End If
            Call bs.Append(v, 8)
            flag = Not flag
        Loop
    End Sub

    Private Function GetModuleMatrix()
        Dim numModulesPerSide
        numModulesPerSide = Module.GetNumModulesPerSide(m_currVersion)

        Dim moduleMatrix()
        ReDim moduleMatrix(numModulesPerSide - 1)

        Dim i
        Dim cols()

        For i = 0 To UBound(moduleMatrix)
            ReDim cols(numModulesPerSide - 1)
            moduleMatrix(i) = cols
        Next

        Call FinderPattern.Place(moduleMatrix)
        Call Separator.Place(moduleMatrix)
        Call TimingPattern.Place(moduleMatrix)

        If m_currVersion >= 2 Then
            Call AlignmentPattern.Place(moduleMatrix, m_currVersion)
        End If

        Call FormatInfo.PlaceTempBlank(moduleMatrix)

        If m_currVersion >= 7 Then
            Call VersionInfo.PlaceTempBlank(moduleMatrix)
        End If

        Call PlaceSymbolChar(moduleMatrix)
        Call RemainderBit.Place(moduleMatrix)

        Dim maskPatternReference
        maskPatternReference = Masking.Apply( _
            moduleMatrix, m_currVersion, m_parent.ErrorCorrectionLevel)

        Call FormatInfo.Place(moduleMatrix, _
                              m_parent.ErrorCorrectionLevel, _
                              maskPatternReference)

        If m_currVersion >= 7 Then
            Call VersionInfo.Place(moduleMatrix, m_currVersion)
        End If

        GetModuleMatrix = moduleMatrix
    End Function

    Private Sub PlaceSymbolChar(ByRef moduleMatrix())
        Dim data
        data = GetEncodingRegionBytes()

        Dim r
        r = UBound(moduleMatrix)

        Dim c
        c = UBound(moduleMatrix(0))

        Dim toLeft
        toLeft = True

        Dim rowDirection
        rowDirection = -1

        Dim bitPos
        Dim i

        For i = 0 To UBound(data)
            bitPos = 7

            Do While bitPos >= 0
                If moduleMatrix(r)(c) = 0 Then
                    If (data(i) And 2 ^ bitPos) > 0 Then
                        moduleMatrix(r)(c) = 1
                    Else
                        moduleMatrix(r)(c) = -1
                    End If

                    bitPos = bitPos - 1
                End If

                If toLeft Then
                    c = c - 1
                Else
                    If (r + rowDirection) < 0 Then
                        r = 0
                        rowDirection = 1
                        c = c - 1

                        If c = 6 Then
                            c = 5
                        End If

                    ElseIf ((r + rowDirection) > UBound(moduleMatrix)) Then
                        r = UBound(moduleMatrix)
                        rowDirection = -1
                        c = c - 1

                        If c = 6 Then
                            c = 5
                        End If

                    Else
                        r = r + rowDirection
                        c = c + 1
                    End If
                End If

                toLeft = Not toLeft
            Loop
        Next
    End Sub

    Public Function Get1bppDIB(ByVal moduleSize, ByVal foreColor, ByVal backColor)
        If moduleSize < 1 Or moduleSize > 31 Then Call Err.Raise(5)

        Dim foreRGB
        foreRGB = ToRGB(foreColor)
        Dim backRGB
        backRGB = ToRGB(backColor)

        If m_dataBitCounter = 0 Then Call Err.Raise(51)

        Dim moduleMatrix
        moduleMatrix = QuietZone.Place(GetModuleMatrix())

        Dim moduleCount
        moduleCount = UBound(moduleMatrix) + 1

        Dim pictWidth
        pictWidth = moduleCount * moduleSize

        Dim pictHeight
        pictHeight = moduleCount * moduleSize

        Dim rowBytesLen
        rowBytesLen = (pictWidth + 7) \ 8

        Dim pack8bit
        If pictWidth Mod 8 > 0 Then
            pack8bit = 8 - (pictWidth Mod 8)
        End If

        Dim pack32bit
        If rowBytesLen Mod 4 > 0 Then
            pack32bit = 8 * (4 - (rowBytesLen Mod 4))
        End If

        Dim rowSize
        rowSize = (pictWidth + pack8bit + pack32bit) \ 8

        Dim bitmapData
        Set bitmapData = New BinaryWriter

        Dim offset
        offset = 0

        Dim bs
        Set bs = New BitSequence

        Dim r, c
        Dim i
        Dim pixelColor

        For r = UBound(moduleMatrix) To 0 Step -1
            Call bs.Clear

            For c = 0 To UBound(moduleMatrix(r))
                If moduleMatrix(r)(c) > 0 Then
                    pixelColor = 0
                Else
                    pixelColor = (2 ^ moduleSize) - 1
                End If

                Call bs.Append(pixelColor, moduleSize)

            Next
            Call bs.Append(0, pack8bit)
            Call bs.Append(0, pack32bit)

            Dim bitmapRow
            bitmapRow = bs.GetBytes()

            For i = 1 To moduleSize
                Call bitmapData.Append(bitmapRow)
            Next
        Next

        Dim ret
        Set ret = Build1bppDIB(bitmapData, pictWidth, pictHeight, foreRGB, backRGB)

        Set Get1bppDIB = ret
    End Function

    Public Function Get24bppDIB(ByVal moduleSize, ByVal foreColor, ByVal backColor)
        If moduleSize < 1 Then Call Err.Raise(5)

        Dim foreRGB
        foreRGB = ToRGB(foreColor)
        Dim backRGB
        backRGB = ToRGB(backColor)

        If m_dataBitCounter = 0 Then Call Err.Raise(51)

        Dim moduleMatrix
        moduleMatrix = QuietZone.Place(GetModuleMatrix())

        Dim pictWidth
        pictWidth = (UBound(moduleMatrix) + 1) * moduleSize

        Dim pictHeight
        pictHeight = pictWidth

        Dim rowBytesLen
        rowBytesLen = 3 * pictWidth

        Dim pack4byte
        If rowBytesLen Mod 4 > 0 Then
            pack4byte = 4 - (rowBytesLen Mod 4)
        End If

        Dim rowSize
        rowSize = rowBytesLen + pack4byte

        Dim bitmapData
        Set bitmapData = New BinaryWriter

        Dim offset
        offset = 0

        Dim r, c
        Dim i

        Dim colorRGB
        Dim bitmapRow()
        Dim idx

        For r = UBound(moduleMatrix) To 0 Step -1
            ReDim bitmapRow(rowSize - 1)
            idx = 0

            For c = 0 To UBound(moduleMatrix(r))
                If moduleMatrix(r)(c) > 0 Then
                    colorRGB = foreRGB
                Else
                    colorRGB = backRGB
                End If

                For i = 1 To moduleSize
                    bitmapRow(idx + 0) = CByte((colorRGB And &HFF0000) \ 2 ^ 16)
                    bitmapRow(idx + 1) = CByte((colorRGB And &HFF00&) \ 2 ^ 8)
                    bitmapRow(idx + 2) = CByte(colorRGB And &HFF&)
                    idx = idx + 3
                Next

            Next

            For i = 1 To pack4byte
                bitmapRow(idx) = CByte(0)
                idx = idx + 1
            Next

            For i = 1 To moduleSize
                Call bitmapData.Append(bitmapRow)
            Next
        Next

        Dim ret
        Set ret = Build24bppDIB(bitmapData, pictWidth, pictHeight)

        Set Get24bppDIB = ret
    End Function

    Public Sub Save1bppDIB(ByVal filePath, ByVal moduleSize, ByVal foreRGB, ByVal backRGB)
        If Len(filePath) = 0 Then Call Err.Raise(5)
        If moduleSize < 1 Or moduleSize > 31 Then Call Err.Raise(5)

        If m_dataBitCounter = 0 Then Call Err.Raise(51)

        Dim dib
        Set dib = Get1bppDIB(moduleSize, foreRGB, backRGB)

        Call dib.SaveToFile(filePath, adSaveCreateOverWrite)
    End Sub

    Public Sub Save24bppDIB(ByVal filePath, ByVal moduleSize, ByVal foreRGB, ByVal backRGB)
        If Len(filePath) = 0 Then Call Err.Raise(5)
        If moduleSize < 1 Then Call Err.Raise(5)

        If m_dataBitCounter = 0 Then Call Err.Raise(51)

        Dim dib
        Set dib = Get24bppDIB(moduleSize, foreRGB, backRGB)

        Call dib.SaveToFile(filePath, adSaveCreateOverWrite)
    End Sub

End Class


Class Symbols

    Private m_items

    Private m_minVersion
    Private m_maxVersion
    Private m_errorCorrectionLevel
    Private m_structuredAppendAllowed
    Private m_byteModeCharsetName

    Private m_structuredAppendParity

    Private m_currSymbol

    Private m_encNum
    Private m_encAlpha
    Private m_encByte
    Private m_encKanji

    Public Sub Init(ByVal ecLevel, ByVal maxVer, ByVal allowStructuredAppend)
        If maxVer < MIN_VERSION Or _
           maxVer > MAX_VERSION Then
            Call Err.Raise(5)
        End If

        Set m_items = New List

        Set m_encNum = CreateEncoder(ENCODINGMODE_NUMERIC)
        Set m_encAlpha = CreateEncoder(ENCODINGMODE_ALPHA_NUMERIC)
        Set m_encByte = CreateEncoder(ENCODINGMODE_EIGHT_BIT_BYTE)
        Set m_encKanji = CreateEncoder(ENCODINGMODE_KANJI)

        m_minVersion = 1
        m_maxVersion = maxVer
        m_errorCorrectionLevel = ecLevel
        m_structuredAppendAllowed = allowStructuredAppend

        m_structuredAppendParity = 0

        Set m_currSymbol = New Symbol
        Call m_currSymbol.Init(Me)
        Call m_items.Add(m_currSymbol)
    End Sub

    Public Property Get Item(ByVal idx)
        Set Item = m_items.Item(idx)
    End Property

    Public Property Get Count()
        Count = m_items.Count
    End Property

    Public Property Get StructuredAppendAllowed()
        StructuredAppendAllowed = m_structuredAppendAllowed
    End Property

    Public Property Get StructuredAppendParity()
        StructuredAppendParity = m_structuredAppendParity
    End Property

    Public Property Get MinVersion()
        MinVersion = m_minVersion
    End Property
    Public Property Let MinVersion(ByVal Value)
        m_minVersion = Value
    End Property

    Public Property Get MaxVersion()
        MaxVersion = m_maxVersion
    End Property

    Public Property Get ErrorCorrectionLevel()
        ErrorCorrectionLevel = m_errorCorrectionLevel
    End Property

    Private Function Add()
        Set m_currSymbol = New Symbol
        Call m_currSymbol.Init(Me)
        Call m_items.Add(m_currSymbol)

        Set Add = m_currSymbol
    End Function

    Public Sub AppendText(ByVal s)
        Dim oldMode
        Dim newMode
        Dim i

        If Len(s) = 0 Then Call Err.Raise(5)

        For i = 1 To Len(s)
            oldMode = m_currSymbol.CurrentEncodingMode

            Select Case oldMode
                Case ENCODINGMODE_UNKNOWN
                    newMode = SelectInitialMode(s, i)
                Case ENCODINGMODE_NUMERIC
                    newMode = SelectModeWhileInNumericMode(s, i)
                Case ENCODINGMODE_ALPHA_NUMERIC
                    newMode = SelectModeWhileInAlphanumericMode(s, i)
                Case ENCODINGMODE_EIGHT_BIT_BYTE
                    newMode = SelectModeWhileInByteMode(s, i)
                Case ENCODINGMODE_KANJI
                    newMode = SelectInitialMode(s, i)
                Case Else
                    Call Err.Raise(51)
            End Select

            If newMode <> oldMode Then
                If Not m_currSymbol.TrySetEncodingMode(newMode, Mid(s, i, 1)) Then
                    If Not m_structuredAppendAllowed Or m_items.Count = 16 Then
                        Call Err.Raise(6)
                    End If

                    Call Add
                    newMode = SelectInitialMode(s, i)
                    Call m_currSymbol.TrySetEncodingMode(newMode, Mid(s, i, 1))
                End If
            End If

            If Not m_currSymbol.TryAppend(Mid(s, i, 1)) Then
                If Not m_structuredAppendAllowed Or m_items.Count = 16 Then
                    Call Err.Raise(6)
                End If

                Call Add
                newMode = SelectInitialMode(s, i)
                Call m_currSymbol.TrySetEncodingMode(newMode, Mid(s, i, 1))
                Call m_currSymbol.TryAppend(Mid(s, i, 1))
            End If
        Next
    End Sub

    Public Sub UpdateParity(ByVal c)
        Dim code
        code = Asc(c) And &HFFFF&

        Dim msb
        Dim lsb

        msb = (code And &HFF00&) \ 2 ^ 8
        lsb = code And &HFF&

        If msb > 0 Then
            m_structuredAppendParity = m_structuredAppendParity Xor msb
        End If

        m_structuredAppendParity = m_structuredAppendParity Xor lsb
    End Sub

    Private Function SelectInitialMode(ByRef s, ByVal startIndex)
        Dim cnt
        Dim flg
        Dim flg1
        Dim flg2
        Dim i

        Dim ver
        ver = m_currSymbol.Version

        If m_encKanji.InSubset(Mid(s, startIndex, 1)) Then
            SelectInitialMode = ENCODINGMODE_KANJI
            Exit Function
        ElseIf m_encByte.InExclusiveSubset(Mid(s, startIndex, 1)) Then
            SelectInitialMode = ENCODINGMODE_EIGHT_BIT_BYTE
            Exit Function
        ElseIf m_encAlpha.InExclusiveSubset(Mid(s, startIndex, 1)) Then
            cnt = 0
            flg = False

            For i = startIndex To Len(s)
                If m_encAlpha.InExclusiveSubset(Mid(s, i, 1)) Then
                    cnt = cnt + 1
                Else
                    Exit For
                End If
            Next

            If 1 <= ver And ver <= 9 Then
                flg = cnt < 6
            ElseIf 10 <= ver And ver <= 26 Then
                flg = cnt < 7
            ElseIf 27 <= ver And ver <= 40 Then
                flg = cnt < 8
            Else
                Call Err.Raise(51)
            End If

            If flg Then
                If (startIndex + cnt) <= Len(s) Then
                    If m_encByte.InExclusiveSubset(Mid(s, startIndex + cnt, 1)) Then
                        SelectInitialMode = ENCODINGMODE_EIGHT_BIT_BYTE
                        Exit Function
                    Else
                        SelectInitialMode = ENCODINGMODE_ALPHA_NUMERIC
                        Exit Function
                    End If
                Else
                    SelectInitialMode = ENCODINGMODE_ALPHA_NUMERIC
                    Exit Function
                End If
            Else
                SelectInitialMode = ENCODINGMODE_ALPHA_NUMERIC
                Exit Function
            End If
        ElseIf m_encNum.InSubset(Mid(s, startIndex, 1)) Then
            cnt = 0
            flg1 = False
            flg2 = False

            For i = startIndex To Len(s)
                If m_encNum.InSubset(Mid(s, i, 1)) Then
                    cnt = cnt + 1
                Else
                    Exit For
                End If
            Next

            If 1 <= ver And ver <= 9 Then
                flg1 = cnt < 4
                flg2 = cnt < 7
            ElseIf 10 <= ver And ver <= 26 Then
                flg1 = cnt < 4
                flg2 = cnt < 8
            ElseIf 27 <= ver And ver <= 40 Then
                flg1 = cnt < 5
                flg2 = cnt < 9
            Else
                Call Err.Raise(51)
             End If

            If flg1 Then
                If (startIndex + cnt) <= Len(s) Then
                    flg1 = m_encByte.InExclusiveSubset(Mid(s, startIndex + cnt, 1))
                Else
                    flg1 = False
                End If
            End If

            If flg2 Then
                If (startIndex + cnt) <= Len(s) Then
                    flg2 = m_encAlpha.InExclusiveSubset(Mid(s, startIndex + cnt, 1))
                Else
                    flg2 = False
                End If
            End If

            If flg1 Then
                SelectInitialMode = ENCODINGMODE_EIGHT_BIT_BYTE
                Exit Function
            ElseIf flg2 Then
                SelectInitialMode = ENCODINGMODE_ALPHA_NUMERIC
                Exit Function
            Else
                SelectInitialMode = ENCODINGMODE_NUMERIC
                Exit Function
            End If
        Else
            Call Err.Raise(51)
        End If
    End Function

    Private Function SelectModeWhileInNumericMode(ByRef s, ByVal startIndex)
        If m_encKanji.InSubset(Mid(s, startIndex, 1)) Then
            SelectModeWhileInNumericMode = ENCODINGMODE_KANJI
            Exit Function
        ElseIf m_encByte.InExclusiveSubset(Mid(s, startIndex, 1)) Then
            SelectModeWhileInNumericMode = ENCODINGMODE_EIGHT_BIT_BYTE
            Exit Function
        ElseIf m_encAlpha.InExclusiveSubset(Mid(s, startIndex, 1)) Then
            SelectModeWhileInNumericMode = ENCODINGMODE_ALPHA_NUMERIC
            Exit Function
        End If

        SelectModeWhileInNumericMode = ENCODINGMODE_NUMERIC
    End Function

    Private Function SelectModeWhileInAlphanumericMode(ByRef s, ByVal startIndex)
        Dim cnt
        Dim flg
        Dim i

        Dim ver
        ver = m_currSymbol.Version

        If m_encKanji.InSubset(Mid(s, startIndex, 1)) Then
            SelectModeWhileInAlphanumericMode = ENCODINGMODE_KANJI
            Exit Function
        ElseIf m_encByte.InExclusiveSubset(Mid(s, startIndex, 1)) Then
            SelectModeWhileInAlphanumericMode = ENCODINGMODE_EIGHT_BIT_BYTE
            Exit Function
        End If

        cnt = 0
        flg = False

        For i = startIndex To Len(s)
            If Not m_encAlpha.InSubset(Mid(s, i, 1)) Then
                Exit For
            End If

            If m_encNum.InSubset(Mid(s, i, 1)) Then
                cnt = cnt + 1
            Else
                flg = True
                Exit For
            End If
        Next

        If flg Then
            If 1 <= ver And ver <= 9 Then
                flg = cnt >= 13
            ElseIf 10 <= ver And ver <= 26 Then
                flg = cnt >= 15
            ElseIf 27 <= ver And ver <= 40 Then
                flg = cnt >= 17
            Else
                Call Err.Raise(51)
            End If

            If flg Then
                SelectModeWhileInAlphanumericMode = ENCODINGMODE_NUMERIC
                Exit Function
            End If
        End If

        SelectModeWhileInAlphanumericMode = ENCODINGMODE_ALPHA_NUMERIC
    End Function

    Private Function SelectModeWhileInByteMode(ByRef s, ByVal startIndex)
        Dim cnt
        Dim flg
        Dim i

        Dim ver
        ver = m_currSymbol.Version

        If m_encKanji.InSubset(Mid(s, startIndex, 1)) Then
            SelectModeWhileInByteMode = ENCODINGMODE_KANJI
            Exit Function
        End If

        For i = startIndex To Len(s)
            If Not m_encByte.InSubset(Mid(s, i, 1)) Then
                Exit For
            End If

            If m_encNum.InSubset(Mid(s, i, 1)) Then
                cnt = cnt + 1
            ElseIf m_encByte.InExclusiveSubset(Mid(s, i, 1)) Then
                flg = True
                Exit For
            Else
                Exit For
            End If
        Next

        If flg Then
            If 1 <= ver And ver <= 9 Then
                flg = cnt >= 6
            ElseIf 10 <= ver And ver <= 26 Then
                flg = cnt >= 8
            ElseIf 27 <= ver And ver <= 40 Then
                flg = cnt >= 9
            Else
                Call Err.Raise(51)
            End If

            If flg Then
                SelectModeWhileInByteMode = ENCODINGMODE_NUMERIC
                Exit Function
            End If
        End If

        cnt = 0
        flg = False

        For i = startIndex To Len(s)
            If Not m_encByte.InSubset(Mid(s, i, 1)) Then
                Exit For
            End If

            If m_encAlpha.InExclusiveSubset(Mid(s, i, 1)) Then
                cnt = cnt + 1
            ElseIf m_encByte.InExclusiveSubset(Mid(s, i, 1)) Then
                flg = True
                Exit For
            Else
                Exit For
            End If

            i = i + 1
        Next

        If flg Then
            If 1 <= ver And ver <= 9 Then
                flg = cnt >= 11
            ElseIf 10 <= ver And ver <= 26 Then
                flg = cnt >= 15
            ElseIf 27 <= ver And ver <= 40 Then
                flg = cnt >= 16
            Else
                Call Err.Raise(51)
            End If

            If flg Then
                SelectModeWhileInByteMode = ENCODINGMODE_ALPHA_NUMERIC
                Exit Function
            End If
        End If

        SelectModeWhileInByteMode = ENCODINGMODE_EIGHT_BIT_BYTE
    End Function

End Class


Class TimingPattern_

    Public Sub Place(ByRef moduleMatrix())
        Dim i
        Dim v

        For i = 8 To UBound(moduleMatrix) - 8
            If i Mod 2 = 0 Then
                v = 2
            Else
                v = -2
            End If

            moduleMatrix(6)(i) = v
            moduleMatrix(i)(6) = v
        Next
    End Sub

End Class


Class VersionInfo_

    Private m_versionInfoValues

    Private Sub Class_Initialize()
        m_versionInfoValues = Array( _
            -1, -1, -1, -1, -1, -1, -1, _
            &H7C94&, &H85BC&, &H9A99&, &HA4D3&, &HBBF6&, &HC762&, &HD847&, &HE60D&, _
            &HF928&, &H10B78, &H1145D, &H12A17, &H13532, &H149A6, &H15683, &H168C9, _
            &H177EC, &H18EC4, &H191E1, &H1AFAB, &H1B08E, &H1CC1A, &H1D33F, &H1ED75, _
            &H1F250, &H209D5, &H216F0, &H228BA, &H2379F, &H24B0B, &H2542E, &H26A64, _
            &H27541, &H28C69 _
        )
    End Sub

    Public Sub Place(ByRef moduleMatrix(), ByVal ver)
        Dim numModulesPerSide
        numModulesPerSide = UBound(moduleMatrix) + 1

        Dim versionInfoValue
        versionInfoValue = m_versionInfoValues(ver)

        Dim p1
        p1 = 0

        Dim p2
        p2 = numModulesPerSide - 11

        Dim i
        Dim v

        For i = 0 To 17
            If (versionInfoValue And 2 ^ i) > 0 Then
                v = 3
            Else
                v = -3
            End If

            moduleMatrix(p1)(p2) = v
            moduleMatrix(p2)(p1) = v

            p2 = p2 + 1

            If i Mod 3 = 2 Then
                p1 = p1 + 1
                p2 = numModulesPerSide - 11
            End If

        Next
    End Sub

    Public Sub PlaceTempBlank(ByRef moduleMatrix())
        Dim numModulesPerSide
        numModulesPerSide = UBound(moduleMatrix) + 1

        Dim i, j

        For i = 0 To 5
            For j = numModulesPerSide - 11 To numModulesPerSide - 9
                moduleMatrix(i)(j) = -3
                moduleMatrix(j)(i) = -3
            Next
        Next
    End Sub

End Class


Class BITMAPFILEHEADER

    Private m_bfType
    Public Property Let bfType(ByVal Value)
        m_bfType = CInt(Value)
    End Property
    Public Property Get bfType()
        bfType = m_bfType
    End Property

    Private m_bfSize
    Public Property Let bfSize(ByVal Value)
        m_bfSize = CLng(Value)
    End Property
    Public Property Get bfSize()
        bfSize = m_bfSize
    End Property

    Private m_bfReserved1
    Public Property Let bfReserved1(ByVal Value)
        m_bfReserved1 = CInt(Value)
    End Property
    Public Property Get bfReserved1()
        bfReserved1 = m_bfReserved1
    End Property

    Private m_bfReserved2
    Public Property Let bfReserved2(ByVal Value)
        m_bfReserved2 = CInt(Value)
    End Property
    Public Property Get bfReserved2()
        bfReserved2 = m_bfReserved2
    End Property

    Private m_bfOffBits
    Public Property Let bfOffBits(ByVal Value)
        m_bfOffBits = CLng(Value)
    End Property
    Public Property Get bfOffBits()
        bfOffBits = m_bfOffBits
    End Property

End Class


Class BITMAPINFOHEADER

    Private m_biSize
    Public Property Let biSize(ByVal Value)
        m_biSize = CLng(Value)
    End Property
    Public Property Get biSize()
        biSize = m_biSize
    End Property

    Private m_biWidth
    Public Property Let biWidth(ByVal Value)
        m_biWidth = CLng(Value)
    End Property
    Public Property Get biWidth()
        biWidth = m_biWidth
    End Property

    Private m_biHeight
    Public Property Let biHeight(ByVal Value)
        m_biHeight = CLng(Value)
    End Property
    Public Property Get biHeight()
        biHeight = m_biHeight
    End Property

    Private m_biPlanes
    Public Property Let biPlanes(ByVal Value)
        m_biPlanes = CInt(Value)
    End Property
    Public Property Get biPlanes()
        biPlanes = m_biPlanes
    End Property

    Private m_biBitCount
    Public Property Let biBitCount(ByVal Value)
        m_biBitCount = CInt(Value)
    End Property
    Public Property Get biBitCount()
        biBitCount = m_biBitCount
    End Property

    Private m_biCompression
    Public Property Let biCompression(ByVal Value)
        m_biCompression = CLng(Value)
    End Property
    Public Property Get biCompression()
        biCompression = m_biCompression
    End Property

    Private m_biSizeImage
    Public Property Let biSizeImage(ByVal Value)
        m_biSizeImage = CLng(Value)
    End Property
    Public Property Get biSizeImage()
        biSizeImage = m_biSizeImage
    End Property

    Private m_biXPelsPerMeter
    Public Property Let biXPelsPerMeter(ByVal Value)
        m_biXPelsPerMeter = CLng(Value)
    End Property
    Public Property Get biXPelsPerMeter()
        biXPelsPerMeter = m_biXPelsPerMeter
    End Property

    Private m_biYPelsPerMeter
    Public Property Let biYPelsPerMeter(ByVal Value)
        m_biYPelsPerMeter = CLng(Value)
    End Property
    Public Property Get biYPelsPerMeter()
        biYPelsPerMeter = m_biYPelsPerMeter
    End Property

    Private m_biClrUsed
    Public Property Let biClrUsed(ByVal Value)
        m_biClrUsed = CLng(Value)
    End Property
    Public Property Get biClrUsed()
        biClrUsed = m_biClrUsed
    End Property

    Private m_biClrImportant
    Public Property Let biClrImportant(ByVal Value)
        m_biClrImportant = CLng(Value)
    End Property
    Public Property Get biClrImportant()
        biClrImportant = m_biClrImportant
    End Property

End Class


Class RGBQUAD

    Private m_rgbBlue
    Public Property Let rgbBlue(ByVal Value)
        m_rgbBlue = CByte(Value)
    End Property
    Public Property Get rgbBlue()
        rgbBlue = m_rgbBlue
    End Property

    Private m_rgbGreen
    Public Property Let rgbGreen(ByVal Value)
        m_rgbGreen = CByte(Value)
    End Property
    Public Property Get rgbGreen()
        rgbGreen = m_rgbGreen
    End Property

    Private m_rgbRed
    Public Property Let rgbRed(ByVal Value)
        m_rgbRed = CByte(Value)
    End Property
    Public Property Get rgbRed()
        rgbRed = m_rgbRed
    End Property

    Private m_rgbReserved
    Public Property Let rgbReserved(ByVal Value)
        m_rgbReserved = CByte(Value)
    End Property
    Public Property Get rgbReserved()
        rgbReserved = m_rgbReserved
    End Property

End Class


Public Sub Main(ByVal args)
    If args.Count = 0 Then Exit Sub
        
    Dim namedArgs
    Set namedArgs = args.Named
    Dim unNamedArgs
    Set unNamedArgs = args.UnNamed

    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts
    Dim data
    Dim inFile
    Dim outFilename

    Dim parentFolder
    parentFolder = fso.getParentFolderName(WScript.ScriptFullName)

    If unNamedArgs.Count > 0 Then
        If fso.FileExists(unNamedArgs(0)) Then
            Set inFile = fso.GetFile(unNamedArgs(0))
            Set ts = inFile.OpenAsTextStream()
            data = ts.ReadAll()
            Call ts.Close
            outFilename = fso.GetParentFolderName(infile.Path) _
            	& "\" & fso.GetBaseName(infile.Name) & ".bmp"
        Else
            Call WScript.Echo("file not found")
            Call WScript.Quit(-1)
        End If
    End If

    Dim foreColor
    foreColor = "#000000"

    Dim backColor
    backColor = "#FFFFFF"

    Dim moduleSize
    moduleSize = 4

    Dim colorDepth
    colorDepth = 24

    Dim ecLevel
    ecLevel = ERRORCORRECTION_LEVEL_M

    Dim temp

    Dim re
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True

    If namedArgs.Count > 0 Then
        If namedArgs.Exists("data") Then
            If Len(data) = 0 Then
                data = namedArgs.Item("data")
            End If
        Else
            If Len(data) = 0 Then
                Call WScript.Echo("argument error ""data""")
                Call WScript.Quit(-1)
            End IF
        End IF

        If namedArgs.Exists("out") Then
            outFilename = namedArgs.Item("out")
        Else
            If Len(outFilename) = 0 Then
                Call WScript.Echo("argument error ""out""")
                Call WScript.Quit(-1)
            End If
        End IF

        re.Pattern = "^#[0-9A-Fa-f]{6}$"

        If namedArgs.Exists("forecolor") Then
            If re.Test(namedArgs.Item("forecolor")) Then
                foreColor = namedArgs.Item("forecolor")
            Else
                Call WScript.Echo("argument error ""forecolor""")
                Call WScript.Quit(-1)
            End If
        End If

        If namedArgs.Exists("backcolor") Then
            If re.Test(namedArgs.Item("backcolor")) Then
                backColor = namedArgs.Item("backcolor")
            Else
                Call WScript.Echo("argument error ""backcolor""")
                Call WScript.Quit(-1)
            End If
        End If

        re.Pattern = "^\d{1,2}$"

        If namedArgs.Exists("modulesize") Then
            If re.Test(namedArgs.Item("modulesize")) Then
                moduleSize = CLng(namedArgs.Item("modulesize"))
            Else
                Call WScript.Echo("argument error ""modulesize""")
                Call WScript.Quit(-1)
            End If

            If moduleSize < 1 Or moduleSize > 31 Then
                Call WScript.Echo("argument error ""modulesize""")
                Call WScript.Quit(-1)
            End If            
        End If

        re.Pattern = "^\d{1,2}$"

        If namedArgs.Exists("colordepth") Then
            If re.Test(namedArgs.Item("colordepth")) Then
                colorDepth = CLng(namedArgs.Item("colordepth"))
            Else
                Call WScript.Echo("argument error ""colordepth""")
                Call WScript.Quit(-1)
            End If

            If colorDepth <> 1 And colorDepth <> 24 Then
                Call WScript.Echo("argument error ""colordepth""")
                Call WScript.Quit(-1)
            End If            
        End If

        re.Pattern = "^[LMQH]$"

        If namedArgs.Exists("ec") Then
            temp = namedArgs.Item("ec")

            If re.Test(temp) Then   
                If UCase(temp) = "L" Then
                    ecLevel = ERRORCORRECTION_LEVEL_L
                ElseIf UCase(temp) = "M" Then
                    ecLevel = ERRORCORRECTION_LEVEL_M
                ElseIf UCase(temp) = "Q" Then
                    ecLevel = ERRORCORRECTION_LEVEL_Q
                ElseIf UCase(temp) = "H" Then
                    ecLevel = ERRORCORRECTION_LEVEL_H
                Else
                    Call WScript.Echo("argument error ""ec""")
                    Call WScript.Quit(-1)
                End If
            Else
                Call WScript.Echo("argument error ""ec""")
                Call WScript.Quit(-1)
            End If
        End If
    End If

    Dim sbls
    Set sbls = CreateSymbols(ecLevel, 40, False)
    Call sbls.AppendText(data)

    If colorDepth = 1 Then
        Call sbls.Item(0).Save1bppDIB(outFilename, moduleSize, foreColor, backColor)
    ElseIf colorDepth = 24 Then
        Call sbls.Item(0).Save24bppDIB(outFilename, moduleSize, foreColor, backColor)
    Else
        Call Err.Raise(51)
    End If

    WScript.Quit(0)
End Sub
