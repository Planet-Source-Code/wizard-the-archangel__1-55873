VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Skipjack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Visual Basic Skipjack Implementation
' David Midkiff (mdj2023@hotmail.com)
' ActiveX Antonio Ramirez Cobos
'
' Standard Skipjack implementation with file support, hex conversion,
' speed string concatenation and overall optimisations for Visual Basic.
' Skipjack is questionable since it was developed by the NSA for use in
' Clipper and Capstone chips ... ciphers from the NSA can suggest backdoors
' or hidden weaknesses but this appears to be a farely decent cipher. I
' would recommend it for medium-risk solutions.
'
' Skipjack is property of the NSA so don't piss them off. :)
'
' Information can be found at:
' http://csrc.nist.gov/encryption/tkencryption.html
'
' {wizard}

'
' Implementation of ObjectControl Interface

Implements IAlgorithm



Event Process(percent As Long)
Event statuschanged(lStatus As Long)

Private m_KeyValue      As String
Private m_F(0 To 255)   As Byte
Private m_Key(0 To 127) As Byte
Private ByteArray()     As Byte
Private hiByte          As Long
Private hiBound         As Long
Private m_status        As Long

Public Function DecryptByte(ByteArray() As Byte, Optional Key As String) As String
    
    Dim i As Long, u As Long, K As Long, temp As Byte, Round As Long, Offset As Long, OrigLen As Long, CipherLen As Long, G(0 To 5) As Byte, NextPercent As Long, CurrPercent As Long, Counter(0 To 32) As Byte, w(0 To 3, 0 To 33) As Integer
    
    
    ' ---------------------------------------------------------------------------
    ' Log info if required
    ' ---------------------------------------------------------------------------
    On Error GoTo errhandler

    If (Len(Key) > 0) Then IAlgorithm_Key = Key
    CipherLen = UBound(ByteArray) + 1
    For Offset = 0 To (CipherLen - 1) Step 2
        temp = ByteArray(Offset)
        ByteArray(Offset) = ByteArray(Offset + 1)
        ByteArray(Offset + 1) = temp
    Next
    For Offset = 0 To (CipherLen - 1) Step 8
        Call CopyMem(w(0, 32), ByteArray(Offset), 8)
        K = 32
        u = 31
        For i = 0 To 32
            Counter(i) = i + 1
        Next
        For Round = 1 To 2
            For i = 1 To 8
                  Call CopyMem(G(4), w(1, K), 2)
                  G(3) = m_F(G(5) Xor m_Key(4 * u + 3)) Xor G(4)
                  G(2) = m_F(G(3) Xor m_Key(4 * u + 2)) Xor G(5)
                  G(0) = m_F(G(2) Xor m_Key(4 * u + 1)) Xor G(3)
                  G(1) = m_F(G(0) Xor m_Key(4 * u)) Xor G(2)
                  Call CopyMem(w(0, K - 1), G(0), 2)
                  w(1, K - 1) = w(0, K - 1) Xor w(2, K) Xor Counter(K - 1)
                  w(2, K - 1) = w(3, K)
                  w(3, K - 1) = w(0, K)
                  u = u - 1
                  K = K - 1
            Next
            For i = 1 To 8
                Call CopyMem(G(4), w(1, K), 2)
                G(3) = m_F(G(5) Xor m_Key(4 * u + 3)) Xor G(4)
                G(2) = m_F(G(3) Xor m_Key(4 * u + 2)) Xor G(5)
                G(0) = m_F(G(2) Xor m_Key(4 * u + 1)) Xor G(3)
                G(1) = m_F(G(0) Xor m_Key(4 * u)) Xor G(2)
                Call CopyMem(w(0, K - 1), G(0), 2)
                w(1, K - 1) = w(2, K)
                w(2, K - 1) = w(3, K)
                w(3, K - 1) = w(0, K) Xor w(1, K) Xor Counter(K - 1)
                u = u - 1
                K = K - 1
            Next
        Next
        w(0, 0) = w(0, 0) Xor w(0, 33)
        w(1, 0) = w(1, 0) Xor w(1, 33)
        w(2, 0) = w(2, 0) Xor w(2, 33)
        w(3, 0) = w(3, 0) Xor w(3, 33)
        Call CopyMem(ByteArray(Offset), w(0, 0), 8)
        If (Offset >= NextPercent) Then
            CurrPercent = Int((Offset / CipherLen) * 100)
            NextPercent = (CipherLen * ((CurrPercent + 1) / 100)) + 1
            RaiseEvent Process(CurrPercent)
        End If
        Call CopyMem(w(0, 33), w(0, 32), 8)
        DoEvents
    Next
    
    For Offset = 0 To (CipherLen - 1) Step 2
        temp = ByteArray(Offset)
        ByteArray(Offset) = ByteArray(Offset + 1)
        ByteArray(Offset + 1) = temp
    Next
    Call CopyMem(OrigLen, ByteArray(8), 4)
    If (CipherLen - OrigLen > 19) Or (CipherLen - OrigLen < 12) Then Call Err.Raise(vbObjectError, , "Incorrect size descriptor in Skipjack decryption")
    Call CopyMem(ByteArray(0), ByteArray(12), OrigLen)
    ReDim Preserve ByteArray(OrigLen - 1)
    If (CurrPercent <> 100) Then RaiseEvent Process(100)

    Exit Function
    
errhandler:

    ' ---------------------------------------------------------------------------
    ' Raise friendly error to the handler
    ' ---------------------------------------------------------------------------
    Call Err.Raise(Err.Number, Err.Source, Err.Description)
End Function

Private Function IAlgorithm_DecryptString(Text As String, Optional Key As String, Optional IsTextInHex As Boolean) As String

    Dim ByteArray() As Byte
    m_status = STAT_BUSY
    RaiseEvent statuschanged(m_status)
    
    On Error GoTo errhandler

    
    If IsTextInHex = True Then Text = DeHex(Text)
    ByteArray() = StrConv(Text, vbFromUnicode)
    Call DecryptByte(ByteArray(), Key)
    IAlgorithm_DecryptString = StrConv(ByteArray(), vbUnicode)
    m_status = STAT_READY
    RaiseEvent statuschanged(m_status)
    Exit Function
    
errhandler:
    m_status = STAT_ERROR
    RaiseEvent statuschanged(m_status)
    ' ---------------------------------------------------------------------------
    ' Raise friendly error to the handler
    ' ---------------------------------------------------------------------------
    Call Err.Raise(Err.Number, Err.Source, Err.Description)
End Function

Public Function EncryptByte(ByteArray() As Byte, Optional Key As String)
    Dim i As Long, K As Long, temp As Byte, Round As Long, Offset As Long, OrigLen As Long, Counter As Long, G(0 To 5) As Byte, CipherLen As Long, NextPercent As Long, CurrPercent As Long, w(0 To 3, 0 To 32) As Integer

    On Error GoTo errhandler

    
    If (Len(Key) > 0) Then IAlgorithm_Key = Key
    OrigLen = UBound(ByteArray) + 1
    CipherLen = OrigLen + 12
    If (CipherLen Mod 8 <> 0) Then CipherLen = CipherLen + 8 - (CipherLen Mod 8)
    ReDim Preserve ByteArray(CipherLen - 1)
    Call CopyMem(ByteArray(12), ByteArray(0), OrigLen)
    Call CopyMem(ByteArray(8), OrigLen, 4)
    Call Randomize
    Call CopyMem(ByteArray(0), CLng(2147483647 * Rnd), 4)
    Call CopyMem(ByteArray(4), CLng(2147483647 * Rnd), 4)
    For Offset = 0 To (CipherLen - 1) Step 2
        temp = ByteArray(Offset)
        ByteArray(Offset) = ByteArray(Offset + 1)
        ByteArray(Offset + 1) = temp
    Next
    For Offset = 0 To (CipherLen - 1) Step 8
        Call CopyMem(w(0, 0), ByteArray(Offset), 8)
        w(0, 0) = w(0, 0) Xor w(0, 32)
        w(1, 0) = w(1, 0) Xor w(1, 32)
        w(2, 0) = w(2, 0) Xor w(2, 32)
        w(3, 0) = w(3, 0) Xor w(3, 32)
        K = 0
        Counter = 1
        For Round = 1 To 2
            For i = 1 To 8
                  Call CopyMem(G(0), w(0, K), 2)
                  G(2) = m_F(G(0) Xor m_Key(4 * K)) Xor G(1)
                  G(3) = m_F(G(2) Xor m_Key(4 * K + 1)) Xor G(0)
                  G(5) = m_F(G(3) Xor m_Key(4 * K + 2)) Xor G(2)
                  G(4) = m_F(G(5) Xor m_Key(4 * K + 3)) Xor G(3)
                  Call CopyMem(w(1, K + 1), G(4), 2)
                  w(0, K + 1) = w(1, K + 1) Xor w(3, K) Xor Counter
                  w(2, K + 1) = w(1, K)
                  w(3, K + 1) = w(2, K)
                  Counter = Counter + 1
                  K = K + 1
            Next
            For i = 1 To 8
                Call CopyMem(G(0), w(0, K), 2)
                G(2) = m_F(G(0) Xor m_Key(4 * K)) Xor G(1)
                G(3) = m_F(G(2) Xor m_Key(4 * K + 1)) Xor G(0)
                G(5) = m_F(G(3) Xor m_Key(4 * K + 2)) Xor G(2)
                G(4) = m_F(G(5) Xor m_Key(4 * K + 3)) Xor G(3)
                Call CopyMem(w(1, K + 1), G(4), 2)
                w(0, K + 1) = w(3, K)
                w(2, K + 1) = w(0, K) Xor w(1, K) Xor Counter
                w(3, K + 1) = w(2, K)
                Counter = Counter + 1
                K = K + 1
            Next
        Next
        Call CopyMem(ByteArray(Offset), w(0, 32), 8)
        If (Offset >= NextPercent) Then
            CurrPercent = Int((Offset / CipherLen) * 100)
            NextPercent = (CipherLen * ((CurrPercent + 1) / 100)) + 1
            RaiseEvent Process(CurrPercent)
        End If
    Next
    For Offset = 0 To (CipherLen - 1) Step 2
        temp = ByteArray(Offset)
        ByteArray(Offset) = ByteArray(Offset + 1)
        ByteArray(Offset + 1) = temp
    Next
    If (CurrPercent <> 100) Then RaiseEvent Process(100)
    
    Exit Function
    
errhandler:

    ' ---------------------------------------------------------------------------
    ' Raise friendly error to the handler
    ' ---------------------------------------------------------------------------
    Call Err.Raise(Err.Number, Err.Source, Err.Description)
End Function

Private Function IAlgorithm_EncryptFile(InFile As String, OutFile As String, Overwrite As Boolean, Optional Key As String) As Boolean

    On Error GoTo errhandler
    m_status = STAT_BUSY
    RaiseEvent statuschanged(m_status)
    
    If FileExist(InFile) = False Then
        IAlgorithm_EncryptFile = False
        Exit Function
    End If
    If FileExist(OutFile) = True And Overwrite = False Then
        IAlgorithm_EncryptFile = False
        Exit Function
    End If
    Dim FileO As Integer, Buffer() As Byte
    FileO = FreeFile
    Open InFile For Binary As #FileO
        ReDim Buffer(0 To LOF(FileO) - 1)
        Get #FileO, , Buffer()
    Close #FileO
    Call EncryptByte(Buffer(), Key)
    If FileExist(OutFile) = True Then Kill OutFile
    FileO = FreeFile
    Open OutFile For Binary As #FileO
        Put #FileO, , Buffer()
    Close #FileO
    IAlgorithm_EncryptFile = True
    m_status = STAT_READY
    RaiseEvent statuschanged(m_status)
    
    Exit Function
    
errhandler:
    m_status = STAT_ERROR
    RaiseEvent statuschanged(m_status)
    ' ---------------------------------------------------------------------------
    ' Raise friendly error to the handler
    ' ---------------------------------------------------------------------------
    Call Err.Raise(Err.Number, Err.Source, Err.Description)
End Function
Private Function IAlgorithm_DecryptFile(InFile As String, OutFile As String, Overwrite As Boolean, Optional Key As String) As Boolean
    m_status = STAT_BUSY
    RaiseEvent statuschanged(m_status)
    ' ---------------------------------------------------------------------------
    ' Log info if required
    ' ---------------------------------------------------------------------------
    On Error GoTo errhandler

    If FileExist(InFile) = False Then
        IAlgorithm_DecryptFile = False
        Exit Function
    End If
    If FileExist(OutFile) = True And Overwrite = False Then
        IAlgorithm_DecryptFile = False
        Exit Function
    End If
    Dim FileO As Integer, Buffer() As Byte
    FileO = FreeFile
    Open InFile For Binary As #FileO
        ReDim Buffer(0 To LOF(FileO) - 1)
        Get #FileO, , Buffer()
    Close #FileO
    Call DecryptByte(Buffer(), Key)
    If FileExist(OutFile) Then Kill OutFile
    
    FileO = FreeFile
    Open OutFile For Binary As #FileO
        Put #FileO, , Buffer()
    Close #FileO
    IAlgorithm_DecryptFile = True
    m_status = STAT_READY
    RaiseEvent statuschanged(m_status)
    
    Exit Function
    
errhandler:
    m_status = STAT_ERROR
    RaiseEvent statuschanged(m_status)
    ' ---------------------------------------------------------------------------
    ' Raise friendly error to the handler
    ' ---------------------------------------------------------------------------
    Call Err.Raise(Err.Number, Err.Source, Err.Description)
End Function

Private Function IAlgorithm_EncryptString(Text As String, Optional Key As String, Optional OutputInHex As Boolean) As String

    Dim ByteArray() As Byte
    m_status = STAT_BUSY
    RaiseEvent statuschanged(m_status)
    
    On Error GoTo errhandler

    
    ByteArray() = StrConv(Text, vbFromUnicode)
    Call EncryptByte(ByteArray(), Key)
    IAlgorithm_EncryptString = StrConv(ByteArray(), vbUnicode)
    If OutputInHex = True Then IAlgorithm_EncryptString = EnHex(IAlgorithm_EncryptString)
    
    m_status = STAT_READY
    RaiseEvent statuschanged(m_status)
    Exit Function
    
errhandler:
    m_status = STAT_ERROR
    RaiseEvent statuschanged(m_status)
    ' ---------------------------------------------------------------------------
    ' Raise friendly error to the handler
    ' ---------------------------------------------------------------------------
    Call Err.Raise(Err.Number, Err.Source, Err.Description)
End Function

Private Property Let IAlgorithm_Key(New_Value As String)
    Dim i As Long, pass() As Byte, PassLen As Long
    
    
    ' ---------------------------------------------------------------------------
    ' Log info if required
    ' ---------------------------------------------------------------------------
    On Error GoTo errhandler
   
    If (New_Value = m_KeyValue) Then Exit Property
    If (Len(New_Value) = 0) Then Exit Property
    PassLen = Len(New_Value)
    pass() = StrConv(New_Value, vbFromUnicode)
    For i = 0 To 127
        m_Key(i) = pass(i Mod PassLen)
    Next
    m_KeyValue = New_Value
    Exit Property
    
errhandler:

    
    ' ---------------------------------------------------------------------------
    ' Raise friendly error to the handler
    ' ---------------------------------------------------------------------------
    Call Err.Raise(Err.Number, Err.Source, Err.Description)
End Property

Private Property Get GData() As String
    Dim StringData As String
    StringData = Space(hiByte)
    CopyMem ByVal StringData, ByVal VarPtr(ByteArray(0)), hiByte
    GData = StringData
End Property
Private Sub Reset()
    hiByte = 0
    hiBound = 1024
    ReDim ByteArray(hiBound)
End Sub
Private Function DeHex(Data As String) As String
    Dim iCount As Double
    Reset
    For iCount = 1 To Len(Data) Step 2
        Append Chr$(Val("&H" & Mid$(Data, iCount, 2)))
    Next
    DeHex = GData
    Reset
End Function

Private Function EnHex(Data As String) As String
    Dim iCount As Double, sTemp As String
    Reset
    For iCount = 1 To Len(Data)
        sTemp = Hex$(Asc(Mid$(Data, iCount, 1)))
        If Len(sTemp) < 2 Then sTemp = "0" & sTemp
        Append sTemp
    Next
    EnHex = GData
    Reset
End Function

Private Function FileExist(Filename As String) As Boolean
    On Error GoTo errorhandler
    Call FileLen(Filename)
    FileExist = True
    Exit Function
    
errorhandler:
    FileExist = False
End Function
Private Sub Append(ByRef StringData As String, Optional Length As Long)
    Dim DataLength As Long
    If Length > 0 Then DataLength = Length Else DataLength = Len(StringData)
    If DataLength + hiByte > hiBound Then
        hiBound = hiBound + 1024
        ReDim Preserve ByteArray(hiBound)
    End If
    CopyMem ByVal VarPtr(ByteArray(hiByte)), ByVal StringData, DataLength
    hiByte = hiByte + DataLength
End Sub

Private Sub Class_Initialize()
    Dim a As Long, Ftable As Variant
    Ftable = Array("A3", "D7", "09", "83", "F8", "48", "F6", "F4", "B3", "21", "15", "78", "99", "B1", "AF", "F9", "E7", "2D", "4D", "8A", "CE", "4C", "CA", "2E", "52", "95", "D9", "1E", "4E", "38", "44", "28", "0A", "DF", "02", "A0", "17", "F1", "60", "68", "12", "B7", "7A", "C3", "E9", "FA", "3D", "53", "96", "84", "6B", "BA", "F2", "63", "9A", "19", "7C", "AE", "E5", "F5", "F7", "16", "6A", "A2", "39", "B6", "7B", "0F", "C1", "93", "81", "1B", "EE", "B4", "1A", "EA", "D0", "91", "2F", "B8", "55", "B9", "DA", "85", "3F", "41", "BF", "E0", "5A", "58", "80", "5F", "66", "0B", "D8", "90", "35", "D5", "C0", "A7", "33", "06", "65", "69", "45", "00", "94", "56", "6D", "98", "9B", "76", "97", "FC", "B2", "C2", "B0", "FE", "DB", "20", "E1", "EB", "D6", "E4", "DD", "47", "4A", "1D", "42", "ED", "9E", "6E", "49", "3C", "CD", "43", "27", "D2", "07", "D4", "DE", "C7", "67", "18", "89", "CB", "30", "1F", "8D", "C6", "8F", "AA", "C8", "74", "DC", "C9", "5D", "5C", "31", "A4", _
                   "70", "88", "61", "2C", "9F", "0D", "2B", "87", "50", "82", "54", "64", "26", "7D", "03", "40", "34", "4B", "1C", "73", "D1", "C4", "FD", "3B", "CC", "FB", "7F", "AB", "E6", "3E", "5B", "A5", "AD", "04", "23", "9C", "14", "51", "22", "F0", "29", "79", "71", "7E", "FF", "8C", "0E", "E2", "0C", "EF", "BC", "72", "75", "6F", "37", "A1", "EC", "D3", "8E", "62", "8B", "86", "10", "E8", "08", "77", "11", "BE", "92", "4F", "24", "C5", "32", "36", "9D", "CF", "F3", "A6", "BB", "AC", "5E", "6C", "A9", "13", "57", "25", "B5", "E3", "BD", "A8", "3A", "01", "05", "59", "2A", "46")

    For a = 0 To 255
          m_F(a) = Val("&H" & Ftable(a))
    Next
    m_status = STAT_READY
End Sub


Private Property Get IAlgorithm_Status() As Long
    IAlgorithm_Status = m_status
End Property

Private Property Let IAlgorithm_Status(lStatus As Long)

    m_status = lStatus
    RaiseEvent statuschanged(lStatus)
 
End Property
Private Function IAlgorithm_ToString() As String
    IAlgorithm_ToString = "Skipjack"
End Function
