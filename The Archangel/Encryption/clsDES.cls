VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Visual Basic Blowfish Implementation
' Algorithm Author: Bruce Schneier
' VB Implementation: David Midkiff (mdj2023@hotmail.com)
' ActiveX Implementation: Antonio Ramirez Cobos
'
' Standard DES implementation with file support, hex conversion,
' speed string concatenation and overall optimisations for Visual Basic.
' Yes I know, DES is nearly dead but it still has it's place in the
' cryptographic community and is good for low-risk security.
'
' Information on the DES cipher can be found at:
' http://csrc.nist.gov/encryption/des
'
'{wizard}
Implements IAlgorithm

Event Process(percent As Long)
Event statuschanged(lStatus As Long)

Private m_Key(0 To 47, 1 To 16) As Byte
Private m_KeyValue              As String
Private m_E(0 To 63)            As Byte
Private m_P(0 To 31)            As Byte
Private m_IP(0 To 63)           As Byte
Private m_PC1(0 To 55)          As Byte
Private m_PC2(0 To 47)          As Byte
Private m_IPInv(0 To 63)        As Byte
Private m_EmptyArray(0 To 63)   As Byte
Private m_LeftShifts(1 To 16)   As Byte
Private m_sBox(0 To 7, 0 To 1, 0 To 1, 0 To 1, 0 To 1, 0 To 1, 0 To 1) As Long

Private ByteArray()             As Byte
Private hiByte                  As Long
Private hiBound                 As Long
Private m_status                As Long

Public Function EncryptByte(ByteArray() As Byte, Optional Key As String)
    Dim a As Long, Offset As Long, OrigLen As Long, CipherLen As Long, CurrPercent As Long, NextPercent As Long, CurrBlock(0 To 7) As Byte, CipherBlock(0 To 7) As Byte
    
    
    ' ---------------------------------------------------------------------------
    ' Log info if required
    ' ---------------------------------------------------------------------------
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
    For Offset = 0 To (CipherLen - 1) Step 8
        Call CopyMem(CurrBlock(0), ByteArray(Offset), 8)
        For a = 0 To 7
            CurrBlock(a) = CurrBlock(a) Xor CipherBlock(a)
        Next
        Call EncryptBlock(CurrBlock())
        Call CopyMem(ByteArray(Offset), CurrBlock(0), 8)
        Call CopyMem(CipherBlock(0), CurrBlock(0), 8)
        If (Offset >= NextPercent) Then
            CurrPercent = Int((Offset / CipherLen) * 100)
            NextPercent = (CipherLen * ((CurrPercent + 1) / 100)) + 1
            RaiseEvent Process(CurrPercent)
        End If
        DoEvents
    Next
    If (CurrPercent <> 100) Then RaiseEvent Process(100)
    
    Exit Function
errhandler:

    ' ---------------------------------------------------------------------------
    ' Raise friendly error to the handler
    ' ---------------------------------------------------------------------------
    Call Err.Raise(Err.Number, Err.Source, Err.Description)
End Function

Public Function DecryptByte(ByteArray() As Byte, Optional Key As String)
    
    Dim a As Long, Offset As Long, OrigLen As Long, CipherLen As Long, CurrPercent As Long, NextPercent As Long, CurrBlock(0 To 7) As Byte, CipherBlock(0 To 7) As Byte
    
    
    
    ' ---------------------------------------------------------------------------
    ' Log info if required
    ' ---------------------------------------------------------------------------
    On Error GoTo errhandler

    If (Len(Key) > 0) Then IAlgorithm_Key = Key
    CipherLen = UBound(ByteArray) + 1
    For Offset = 0 To (CipherLen - 1) Step 8
        Call CopyMem(CurrBlock(0), ByteArray(Offset), 8)
        Call DecryptBlock(CurrBlock())
        For a = 0 To 7
            CurrBlock(a) = CurrBlock(a) Xor CipherBlock(a)
        Next
        Call CopyMem(CipherBlock(0), ByteArray(Offset), 8)
        Call CopyMem(ByteArray(Offset), CurrBlock(0), 8)
        If (Offset >= NextPercent) Then
            CurrPercent = Int((Offset / CipherLen) * 100)
            NextPercent = (CipherLen * ((CurrPercent + 1) / 100)) + 1
            RaiseEvent Process(CurrPercent)
        End If
        DoEvents
    Next
    Call CopyMem(OrigLen, ByteArray(8), 4)
    If (CipherLen - OrigLen > 19) Or (CipherLen - OrigLen < 12) Then Call Err.Raise(vbObjectError, , "Incorrect size descriptor in DES decryption")
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


Private Property Let IAlgorithm_Key(New_Value As String)

Dim a As Long, i As Long, C(0 To 27) As Byte, D(0 To 27) As Byte, K(0 To 55) As Byte, CD(0 To 55) As Byte, temp(0 To 1) As Byte, KeyBin(0 To 63) As Byte, KeySchedule(0 To 63) As Byte

On Error GoTo errhandler


    If (m_KeyValue = New_Value) Then Exit Property
    m_KeyValue = New_Value
    Call Byte2Bin(StrConv(New_Value, vbFromUnicode), IIf(Len(New_Value) > 8, 8, Len(New_Value)), KeyBin())
    For a = 0 To 55
        KeySchedule(a) = KeyBin(m_PC1(a))
    Next
    Call CopyMem(C(0), KeySchedule(0), 28)
    Call CopyMem(D(0), KeySchedule(28), 28)
    
    For i = 1 To 16
        Call CopyMem(temp(0), C(0), m_LeftShifts(i))
        Call CopyMem(C(0), C(m_LeftShifts(i)), 28 - m_LeftShifts(i))
        Call CopyMem(C(28 - m_LeftShifts(i)), temp(0), m_LeftShifts(i))
        Call CopyMem(temp(0), D(0), m_LeftShifts(i))
        Call CopyMem(D(0), D(m_LeftShifts(i)), 28 - m_LeftShifts(i))
        Call CopyMem(D(28 - m_LeftShifts(i)), temp(0), m_LeftShifts(i))
        Call CopyMem(CD(0), C(0), 28)
        Call CopyMem(CD(28), D(0), 28)
        For a = 0 To 47
            m_Key(a, i) = CD(m_PC2(a))
        Next
    Next
Exit Property
errhandler:
    ' ---------------------------------------------------------------------------
    ' Raise friendly error to the handler
    ' ---------------------------------------------------------------------------
    Call Err.Raise(Err.Number, Err.Source, Err.Description)
End Property
Private Static Sub Byte2Bin(ByteArray() As Byte, ByteLen As Long, BinaryArray() As Byte)
    Dim a As Long, ByteValue As Byte, BinLength As Long
    Call CopyMem(BinaryArray(0), m_EmptyArray(0), ByteLen * 8)
    BinLength = 0
    For a = 0 To (ByteLen - 1)
        ByteValue = ByteArray(a)
        If (ByteValue And 128) Then BinaryArray(BinLength) = 1
        If (ByteValue And 64) Then BinaryArray(BinLength + 1) = 1
        If (ByteValue And 32) Then BinaryArray(BinLength + 2) = 1
        If (ByteValue And 16) Then BinaryArray(BinLength + 3) = 1
        If (ByteValue And 8) Then BinaryArray(BinLength + 4) = 1
        If (ByteValue And 4) Then BinaryArray(BinLength + 5) = 1
        If (ByteValue And 2) Then BinaryArray(BinLength + 6) = 1
        If (ByteValue And 1) Then BinaryArray(BinLength + 7) = 1
        BinLength = BinLength + 8
    Next
End Sub
Private Static Sub Bin2Byte(BinaryArray() As Byte, ByteLen As Long, ByteArray() As Byte)
    Dim a As Long, ByteValue As Byte, BinLength As Long
    BinLength = 0
    For a = 0 To (ByteLen - 1)
        ByteValue = 0
        If (BinaryArray(BinLength) = 1) Then ByteValue = ByteValue + 128
        If (BinaryArray(BinLength + 1) = 1) Then ByteValue = ByteValue + 64
        If (BinaryArray(BinLength + 2) = 1) Then ByteValue = ByteValue + 32
        If (BinaryArray(BinLength + 3) = 1) Then ByteValue = ByteValue + 16
        If (BinaryArray(BinLength + 4) = 1) Then ByteValue = ByteValue + 8
        If (BinaryArray(BinLength + 5) = 1) Then ByteValue = ByteValue + 4
        If (BinaryArray(BinLength + 6) = 1) Then ByteValue = ByteValue + 2
        If (BinaryArray(BinLength + 7) = 1) Then ByteValue = ByteValue + 1
        ByteArray(a) = ByteValue
        BinLength = BinLength + 8
    Next
End Sub
Private Static Sub EncryptBlock(BlockData() As Byte)
    Dim a As Long, i As Long, L(0 To 31) As Byte, R(0 To 31) As Byte, RL(0 To 63) As Byte, sBox(0 To 31) As Byte, LiRi(0 To 31) As Byte, ERxorK(0 To 47) As Byte, BinBlock(0 To 63) As Byte
    
    Call Byte2Bin(BlockData(), 8, BinBlock())
    For a = 0 To 31
        L(a) = BinBlock(m_IP(a))
        R(a) = BinBlock(m_IP(a + 32))
    Next
    For i = 1 To 16
        ERxorK(0) = R(31) Xor m_Key(0, i)
        ERxorK(1) = R(0) Xor m_Key(1, i)
        ERxorK(2) = R(1) Xor m_Key(2, i)
        ERxorK(3) = R(2) Xor m_Key(3, i)
        ERxorK(4) = R(3) Xor m_Key(4, i)
        ERxorK(5) = R(4) Xor m_Key(5, i)
        ERxorK(6) = R(3) Xor m_Key(6, i)
        ERxorK(7) = R(4) Xor m_Key(7, i)
        ERxorK(8) = R(5) Xor m_Key(8, i)
        ERxorK(9) = R(6) Xor m_Key(9, i)
        ERxorK(10) = R(7) Xor m_Key(10, i)
        ERxorK(11) = R(8) Xor m_Key(11, i)
        ERxorK(12) = R(7) Xor m_Key(12, i)
        ERxorK(13) = R(8) Xor m_Key(13, i)
        ERxorK(14) = R(9) Xor m_Key(14, i)
        ERxorK(15) = R(10) Xor m_Key(15, i)
        ERxorK(16) = R(11) Xor m_Key(16, i)
        ERxorK(17) = R(12) Xor m_Key(17, i)
        ERxorK(18) = R(11) Xor m_Key(18, i)
        ERxorK(19) = R(12) Xor m_Key(19, i)
        ERxorK(20) = R(13) Xor m_Key(20, i)
        ERxorK(21) = R(14) Xor m_Key(21, i)
        ERxorK(22) = R(15) Xor m_Key(22, i)
        ERxorK(23) = R(16) Xor m_Key(23, i)
        ERxorK(24) = R(15) Xor m_Key(24, i)
        ERxorK(25) = R(16) Xor m_Key(25, i)
        ERxorK(26) = R(17) Xor m_Key(26, i)
        ERxorK(27) = R(18) Xor m_Key(27, i)
        ERxorK(28) = R(19) Xor m_Key(28, i)
        ERxorK(29) = R(20) Xor m_Key(29, i)
        ERxorK(30) = R(19) Xor m_Key(30, i)
        ERxorK(31) = R(20) Xor m_Key(31, i)
        ERxorK(32) = R(21) Xor m_Key(32, i)
        ERxorK(33) = R(22) Xor m_Key(33, i)
        ERxorK(34) = R(23) Xor m_Key(34, i)
        ERxorK(35) = R(24) Xor m_Key(35, i)
        ERxorK(36) = R(23) Xor m_Key(36, i)
        ERxorK(37) = R(24) Xor m_Key(37, i)
        ERxorK(38) = R(25) Xor m_Key(38, i)
        ERxorK(39) = R(26) Xor m_Key(39, i)
        ERxorK(40) = R(27) Xor m_Key(40, i)
        ERxorK(41) = R(28) Xor m_Key(41, i)
        ERxorK(42) = R(27) Xor m_Key(42, i)
        ERxorK(43) = R(28) Xor m_Key(43, i)
        ERxorK(44) = R(29) Xor m_Key(44, i)
        ERxorK(45) = R(30) Xor m_Key(45, i)
        ERxorK(46) = R(31) Xor m_Key(46, i)
        ERxorK(47) = R(0) Xor m_Key(47, i)
        Call CopyMem(sBox(0), m_sBox(0, ERxorK(0), ERxorK(1), ERxorK(2), ERxorK(3), ERxorK(4), ERxorK(5)), 4)
        Call CopyMem(sBox(4), m_sBox(1, ERxorK(6), ERxorK(7), ERxorK(8), ERxorK(9), ERxorK(10), ERxorK(11)), 4)
        Call CopyMem(sBox(8), m_sBox(2, ERxorK(12), ERxorK(13), ERxorK(14), ERxorK(15), ERxorK(16), ERxorK(17)), 4)
        Call CopyMem(sBox(12), m_sBox(3, ERxorK(18), ERxorK(19), ERxorK(20), ERxorK(21), ERxorK(22), ERxorK(23)), 4)
        Call CopyMem(sBox(16), m_sBox(4, ERxorK(24), ERxorK(25), ERxorK(26), ERxorK(27), ERxorK(28), ERxorK(29)), 4)
        Call CopyMem(sBox(20), m_sBox(5, ERxorK(30), ERxorK(31), ERxorK(32), ERxorK(33), ERxorK(34), ERxorK(35)), 4)
        Call CopyMem(sBox(24), m_sBox(6, ERxorK(36), ERxorK(37), ERxorK(38), ERxorK(39), ERxorK(40), ERxorK(41)), 4)
        Call CopyMem(sBox(28), m_sBox(7, ERxorK(42), ERxorK(43), ERxorK(44), ERxorK(45), ERxorK(46), ERxorK(47)), 4)
        LiRi(0) = L(0) Xor sBox(15)
        LiRi(1) = L(1) Xor sBox(6)
        LiRi(2) = L(2) Xor sBox(19)
        LiRi(3) = L(3) Xor sBox(20)
        LiRi(4) = L(4) Xor sBox(28)
        LiRi(5) = L(5) Xor sBox(11)
        LiRi(6) = L(6) Xor sBox(27)
        LiRi(7) = L(7) Xor sBox(16)
        LiRi(8) = L(8) Xor sBox(0)
        LiRi(9) = L(9) Xor sBox(14)
        LiRi(10) = L(10) Xor sBox(22)
        LiRi(11) = L(11) Xor sBox(25)
        LiRi(12) = L(12) Xor sBox(4)
        LiRi(13) = L(13) Xor sBox(17)
        LiRi(14) = L(14) Xor sBox(30)
        LiRi(15) = L(15) Xor sBox(9)
        LiRi(16) = L(16) Xor sBox(1)
        LiRi(17) = L(17) Xor sBox(7)
        LiRi(18) = L(18) Xor sBox(23)
        LiRi(19) = L(19) Xor sBox(13)
        LiRi(20) = L(20) Xor sBox(31)
        LiRi(21) = L(21) Xor sBox(26)
        LiRi(22) = L(22) Xor sBox(2)
        LiRi(23) = L(23) Xor sBox(8)
        LiRi(24) = L(24) Xor sBox(18)
        LiRi(25) = L(25) Xor sBox(12)
        LiRi(26) = L(26) Xor sBox(29)
        LiRi(27) = L(27) Xor sBox(5)
        LiRi(28) = L(28) Xor sBox(21)
        LiRi(29) = L(29) Xor sBox(10)
        LiRi(30) = L(30) Xor sBox(3)
        LiRi(31) = L(31) Xor sBox(24)
        Call CopyMem(L(0), R(0), 32)
        Call CopyMem(R(0), LiRi(0), 32)
    Next
    Call CopyMem(RL(0), R(0), 32)
    Call CopyMem(RL(32), L(0), 32)
    For a = 0 To 63
        BinBlock(a) = RL(m_IPInv(a))
    Next
    Call Bin2Byte(BinBlock(), 8, BlockData())
End Sub
Private Static Sub DecryptBlock(BlockData() As Byte)
    Dim a As Long, i As Long, L(0 To 31) As Byte, R(0 To 31) As Byte, RL(0 To 63) As Byte, sBox(0 To 31) As Byte, LiRi(0 To 31) As Byte, ERxorK(0 To 47) As Byte, BinBlock(0 To 63) As Byte
    Call Byte2Bin(BlockData(), 8, BinBlock())
    For a = 0 To 31
        L(a) = BinBlock(m_IP(a))
        R(a) = BinBlock(m_IP(a + 32))
    Next
    For i = 16 To 1 Step -1
        ERxorK(0) = R(31) Xor m_Key(0, i)
        ERxorK(1) = R(0) Xor m_Key(1, i)
        ERxorK(2) = R(1) Xor m_Key(2, i)
        ERxorK(3) = R(2) Xor m_Key(3, i)
        ERxorK(4) = R(3) Xor m_Key(4, i)
        ERxorK(5) = R(4) Xor m_Key(5, i)
        ERxorK(6) = R(3) Xor m_Key(6, i)
        ERxorK(7) = R(4) Xor m_Key(7, i)
        ERxorK(8) = R(5) Xor m_Key(8, i)
        ERxorK(9) = R(6) Xor m_Key(9, i)
        ERxorK(10) = R(7) Xor m_Key(10, i)
        ERxorK(11) = R(8) Xor m_Key(11, i)
        ERxorK(12) = R(7) Xor m_Key(12, i)
        ERxorK(13) = R(8) Xor m_Key(13, i)
        ERxorK(14) = R(9) Xor m_Key(14, i)
        ERxorK(15) = R(10) Xor m_Key(15, i)
        ERxorK(16) = R(11) Xor m_Key(16, i)
        ERxorK(17) = R(12) Xor m_Key(17, i)
        ERxorK(18) = R(11) Xor m_Key(18, i)
        ERxorK(19) = R(12) Xor m_Key(19, i)
        ERxorK(20) = R(13) Xor m_Key(20, i)
        ERxorK(21) = R(14) Xor m_Key(21, i)
        ERxorK(22) = R(15) Xor m_Key(22, i)
        ERxorK(23) = R(16) Xor m_Key(23, i)
        ERxorK(24) = R(15) Xor m_Key(24, i)
        ERxorK(25) = R(16) Xor m_Key(25, i)
        ERxorK(26) = R(17) Xor m_Key(26, i)
        ERxorK(27) = R(18) Xor m_Key(27, i)
        ERxorK(28) = R(19) Xor m_Key(28, i)
        ERxorK(29) = R(20) Xor m_Key(29, i)
        ERxorK(30) = R(19) Xor m_Key(30, i)
        ERxorK(31) = R(20) Xor m_Key(31, i)
        ERxorK(32) = R(21) Xor m_Key(32, i)
        ERxorK(33) = R(22) Xor m_Key(33, i)
        ERxorK(34) = R(23) Xor m_Key(34, i)
        ERxorK(35) = R(24) Xor m_Key(35, i)
        ERxorK(36) = R(23) Xor m_Key(36, i)
        ERxorK(37) = R(24) Xor m_Key(37, i)
        ERxorK(38) = R(25) Xor m_Key(38, i)
        ERxorK(39) = R(26) Xor m_Key(39, i)
        ERxorK(40) = R(27) Xor m_Key(40, i)
        ERxorK(41) = R(28) Xor m_Key(41, i)
        ERxorK(42) = R(27) Xor m_Key(42, i)
        ERxorK(43) = R(28) Xor m_Key(43, i)
        ERxorK(44) = R(29) Xor m_Key(44, i)
        ERxorK(45) = R(30) Xor m_Key(45, i)
        ERxorK(46) = R(31) Xor m_Key(46, i)
        ERxorK(47) = R(0) Xor m_Key(47, i)
        Call CopyMem(sBox(0), m_sBox(0, ERxorK(0), ERxorK(1), ERxorK(2), ERxorK(3), ERxorK(4), ERxorK(5)), 4)
        Call CopyMem(sBox(4), m_sBox(1, ERxorK(6), ERxorK(7), ERxorK(8), ERxorK(9), ERxorK(10), ERxorK(11)), 4)
        Call CopyMem(sBox(8), m_sBox(2, ERxorK(12), ERxorK(13), ERxorK(14), ERxorK(15), ERxorK(16), ERxorK(17)), 4)
        Call CopyMem(sBox(12), m_sBox(3, ERxorK(18), ERxorK(19), ERxorK(20), ERxorK(21), ERxorK(22), ERxorK(23)), 4)
        Call CopyMem(sBox(16), m_sBox(4, ERxorK(24), ERxorK(25), ERxorK(26), ERxorK(27), ERxorK(28), ERxorK(29)), 4)
        Call CopyMem(sBox(20), m_sBox(5, ERxorK(30), ERxorK(31), ERxorK(32), ERxorK(33), ERxorK(34), ERxorK(35)), 4)
        Call CopyMem(sBox(24), m_sBox(6, ERxorK(36), ERxorK(37), ERxorK(38), ERxorK(39), ERxorK(40), ERxorK(41)), 4)
        Call CopyMem(sBox(28), m_sBox(7, ERxorK(42), ERxorK(43), ERxorK(44), ERxorK(45), ERxorK(46), ERxorK(47)), 4)
        LiRi(0) = L(0) Xor sBox(15)
        LiRi(1) = L(1) Xor sBox(6)
        LiRi(2) = L(2) Xor sBox(19)
        LiRi(3) = L(3) Xor sBox(20)
        LiRi(4) = L(4) Xor sBox(28)
        LiRi(5) = L(5) Xor sBox(11)
        LiRi(6) = L(6) Xor sBox(27)
        LiRi(7) = L(7) Xor sBox(16)
        LiRi(8) = L(8) Xor sBox(0)
        LiRi(9) = L(9) Xor sBox(14)
        LiRi(10) = L(10) Xor sBox(22)
        LiRi(11) = L(11) Xor sBox(25)
        LiRi(12) = L(12) Xor sBox(4)
        LiRi(13) = L(13) Xor sBox(17)
        LiRi(14) = L(14) Xor sBox(30)
        LiRi(15) = L(15) Xor sBox(9)
        LiRi(16) = L(16) Xor sBox(1)
        LiRi(17) = L(17) Xor sBox(7)
        LiRi(18) = L(18) Xor sBox(23)
        LiRi(19) = L(19) Xor sBox(13)
        LiRi(20) = L(20) Xor sBox(31)
        LiRi(21) = L(21) Xor sBox(26)
        LiRi(22) = L(22) Xor sBox(2)
        LiRi(23) = L(23) Xor sBox(8)
        LiRi(24) = L(24) Xor sBox(18)
        LiRi(25) = L(25) Xor sBox(12)
        LiRi(26) = L(26) Xor sBox(29)
        LiRi(27) = L(27) Xor sBox(5)
        LiRi(28) = L(28) Xor sBox(21)
        LiRi(29) = L(29) Xor sBox(10)
        LiRi(30) = L(30) Xor sBox(3)
        LiRi(31) = L(31) Xor sBox(24)
        Call CopyMem(L(0), R(0), 32)
        Call CopyMem(R(0), LiRi(0), 32)
    Next
    Call CopyMem(RL(0), R(0), 32)
    Call CopyMem(RL(32), L(0), 32)
    For a = 0 To 63
        BinBlock(a) = RL(m_IPInv(a))
    Next
    Call Bin2Byte(BinBlock(), 8, BlockData())
End Sub

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


Private Sub Class_Initialize()
    Dim i As Long, vE As Variant, vP As Variant, vIP As Variant, vPC1 As Variant, vPC2 As Variant, vIPInv As Variant, vSbox(0 To 7) As Variant
    
    vIP = Array(58, 50, 42, 34, 26, 18, 10, 2, 60, 52, 44, 36, 28, 20, 12, 4, 62, 54, 46, 38, 30, 22, 14, 6, 64, 56, 48, 40, 32, 24, 16, 8, 57, 49, 41, 33, 25, 17, 9, 1, 59, 51, 43, 35, 27, 19, 11, 3, 61, 53, 45, 37, 29, 21, 13, 5, 63, 55, 47, 39, 31, 23, 15, 7)
    For i = LBound(vIP) To UBound(vIP)
        m_IP(i) = (vIP(i) - 1)
    Next
    vE = Array(32, 1, 2, 3, 4, 5, 4, 5, 6, 7, 8, 9, 8, 9, 10, 11, 12, 13, 12, 13, 14, 15, 16, 17, 16, 17, 18, 19, 20, 21, 20, 21, 22, 23, 24, 25, 24, 25, 26, 27, 28, 29, 28, 29, 30, 31, 32, 1)
    For i = LBound(vE) To UBound(vE)
        m_E(i) = (vE(i) - 1)
    Next
    vPC1 = Array(57, 49, 41, 33, 25, 17, 9, 1, 58, 50, 42, 34, 26, 18, 10, 2, 59, 51, 43, 35, 27, 19, 11, 3, 60, 52, 44, 36, 63, 55, 47, 39, 31, 23, 15, 7, 62, 54, 46, 38, 30, 22, 14, 6, 61, 53, 45, 37, 29, 21, 13, 5, 28, 20, 12, 4)
    For i = LBound(vPC1) To UBound(vPC1)
        m_PC1(i) = (vPC1(i) - 1)
    Next
    vPC2 = Array(14, 17, 11, 24, 1, 5, 3, 28, 15, 6, 21, 10, 23, 19, 12, 4, 26, 8, 16, 7, 27, 20, 13, 2, 41, 52, 31, 37, 47, 55, 30, 40, 51, 45, 33, 48, 44, 49, 39, 56, 34, 53, 46, 42, 50, 36, 29, 32)
    For i = LBound(vPC2) To UBound(vPC2)
        m_PC2(i) = (vPC2(i) - 1)
    Next
    vIPInv = Array(40, 8, 48, 16, 56, 24, 64, 32, 39, 7, 47, 15, 55, 23, 63, 31, 38, 6, 46, 14, 54, 22, 62, 30, 37, 5, 45, 13, 53, 21, 61, 29, 36, 4, 44, 12, 52, 20, 60, 28, 35, 3, 43, 11, 51, 19, 59, 27, 34, 2, 42, 10, 50, 18, 58, 26, 33, 1, 41, 9, 49, 17, 57, 25)
    For i = LBound(vIPInv) To UBound(vIPInv)
        m_IPInv(i) = (vIPInv(i) - 1)
    Next
    vP = Array(16, 7, 20, 21, 29, 12, 28, 17, 1, 15, 23, 26, 5, 18, 31, 10, 2, 8, 24, 14, 32, 27, 3, 9, 19, 13, 30, 6, 22, 11, 4, 25)
    For i = LBound(vP) To UBound(vP)
        m_P(i) = (vP(i) - 1)
    Next
    For i = 1 To 16
        Select Case i
        Case 1, 2, 9, 16
            m_LeftShifts(i) = 1
        Case Else
            m_LeftShifts(i) = 2
        End Select
    Next
    vSbox(0) = Array(14, 4, 13, 1, 2, 15, 11, 8, 3, 10, 6, 12, 5, 9, 0, 7, 0, 15, 7, 4, 14, 2, 13, 1, 10, 6, 12, 11, 9, 5, 3, 8, 4, 1, 14, 8, 13, 6, 2, 11, 15, 12, 9, 7, 3, 10, 5, 0, 15, 12, 8, 2, 4, 9, 1, 7, 5, 11, 3, 14, 10, 0, 6, 13)
    vSbox(1) = Array(15, 1, 8, 14, 6, 11, 3, 4, 9, 7, 2, 13, 12, 0, 5, 10, 3, 13, 4, 7, 15, 2, 8, 14, 12, 0, 1, 10, 6, 9, 11, 5, 0, 14, 7, 11, 10, 4, 13, 1, 5, 8, 12, 6, 9, 3, 2, 15, 13, 8, 10, 1, 3, 15, 4, 2, 11, 6, 7, 12, 0, 5, 14, 9)
    vSbox(2) = Array(10, 0, 9, 14, 6, 3, 15, 5, 1, 13, 12, 7, 11, 4, 2, 8, 13, 7, 0, 9, 3, 4, 6, 10, 2, 8, 5, 14, 12, 11, 15, 1, 13, 6, 4, 9, 8, 15, 3, 0, 11, 1, 2, 12, 5, 10, 14, 7, 1, 10, 13, 0, 6, 9, 8, 7, 4, 15, 14, 3, 11, 5, 2, 12)
    vSbox(3) = Array(7, 13, 14, 3, 0, 6, 9, 10, 1, 2, 8, 5, 11, 12, 4, 15, 13, 8, 11, 5, 6, 15, 0, 3, 4, 7, 2, 12, 1, 10, 14, 9, 10, 6, 9, 0, 12, 11, 7, 13, 15, 1, 3, 14, 5, 2, 8, 4, 3, 15, 0, 6, 10, 1, 13, 8, 9, 4, 5, 11, 12, 7, 2, 14)
    vSbox(4) = Array(2, 12, 4, 1, 7, 10, 11, 6, 8, 5, 3, 15, 13, 0, 14, 9, 14, 11, 2, 12, 4, 7, 13, 1, 5, 0, 15, 10, 3, 9, 8, 6, 4, 2, 1, 11, 10, 13, 7, 8, 15, 9, 12, 5, 6, 3, 0, 14, 11, 8, 12, 7, 1, 14, 2, 13, 6, 15, 0, 9, 10, 4, 5, 3)
    vSbox(5) = Array(12, 1, 10, 15, 9, 2, 6, 8, 0, 13, 3, 4, 14, 7, 5, 11, 10, 15, 4, 2, 7, 12, 9, 5, 6, 1, 13, 14, 0, 11, 3, 8, 9, 14, 15, 5, 2, 8, 12, 3, 7, 0, 4, 10, 1, 13, 11, 6, 4, 3, 2, 12, 9, 5, 15, 10, 11, 14, 1, 7, 6, 0, 8, 13)
    vSbox(6) = Array(4, 11, 2, 14, 15, 0, 8, 13, 3, 12, 9, 7, 5, 10, 6, 1, 13, 0, 11, 7, 4, 9, 1, 10, 14, 3, 5, 12, 2, 15, 8, 6, 1, 4, 11, 13, 12, 3, 7, 14, 10, 15, 6, 8, 0, 5, 9, 2, 6, 11, 13, 8, 1, 4, 10, 7, 9, 5, 0, 15, 14, 2, 3, 12)
    vSbox(7) = Array(13, 2, 8, 4, 6, 15, 11, 1, 10, 9, 3, 14, 5, 0, 12, 7, 1, 15, 13, 8, 10, 3, 7, 4, 12, 5, 6, 11, 0, 14, 9, 2, 7, 11, 4, 1, 9, 12, 14, 2, 0, 6, 10, 13, 15, 3, 5, 8, 2, 1, 14, 7, 4, 10, 8, 13, 15, 12, 9, 0, 3, 5, 6, 11)
    
    Dim lBox As Long, lRow As Long, lColumn As Long, TheByte(0) As Byte, TheBin(0 To 7) As Byte, a As Byte, B As Byte, C As Byte, D As Byte, E As Byte, F As Byte
    For lBox = 0 To 7
        For a = 0 To 1
            For B = 0 To 1
                For C = 0 To 1
                    For D = 0 To 1
                        For E = 0 To 1
                            For F = 0 To 1
                                lRow = a * 2 + F
                                lColumn = B * 8 + C * 4 + D * 2 + E
                                TheByte(0) = vSbox(lBox)(lRow * 16 + lColumn)
                                Call Byte2Bin(TheByte(), 1, TheBin())
                                Call CopyMem(m_sBox(lBox, a, B, C, D, E, F), TheBin(4), 4)
                            Next
                        Next
                    Next
                Next
            Next
        Next
    Next
    m_status = STAT_READY
End Sub

Private Function IAlgorithm_ToString() As String
    IAlgorithm_ToString = "DES"
End Function


Private Property Get IAlgorithm_Status() As Long
    IAlgorithm_Status = m_status
End Property

Private Property Let IAlgorithm_Status(lStatus As Long)

    m_status = lStatus
    RaiseEvent statuschanged(lStatus)
 
End Property
