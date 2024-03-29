VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Twofish"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Visual Basic Twofish Implementation
' Algorithm Author: Bruce Schneier
' VB Implementation: David Midkiff (mdj2023@hotmail.com)
' ActiveX Implementation: Antonio Ramirez Cobos
'
' Standard Twofish implementation with file support, hex conversion,
' speed string concatenation and overall optimisations for Visual Basic.
' Twofish is simply another powerful encryption algorithm from Bruce Schneier
' that was a finalist for AES. Export restrictions have been lifted on it
' and I would recommend this as another candidate for high risk level
' solutions.
'
' Information can be found at:
' http://www.counterpane.com/twofish.html
'
'{wizard}

Implements IAlgorithm


Event Process(percent As Long)
Event statuschanged(lStatus As Long)

Public Enum TWOFISHKEYLENGTH
    TWOFISH_256 = 256
    TWOFISH_196 = 196
    TWOFISH_128 = 128
    TWOFISH_64 = 64
End Enum

Private Const ROUNDS = 16
Private Const BLOCK_SIZE = 16
Private Const MAX_ROUNDS = 16
Private Const INPUT_WHITEN = 0
Private Const OUTPUT_WHITEN = INPUT_WHITEN + BLOCK_SIZE / 4
Private Const ROUND_SUBKEYS = OUTPUT_WHITEN + BLOCK_SIZE / 4
Private Const GF256_FDBK_2 = &H169 / 2
Private Const GF256_FDBK_4 = &H169 / 4
Private MDS(0 To 3, 0 To 255)   As Long
Private p(0 To 1, 0 To 255)     As Byte
Private m_RunningCompiled       As Boolean
Private sBox(0 To 1023)         As Long
Private sKey()                  As Long
Private ByteArray()             As Byte
Private hiByte                  As Long
Private hiBound                 As Long
Private m_status                As Long

Private Property Let IAlgorithm_Key(New_Value As String)
  
    On Error GoTo errhandler
    Dim KeyLength As Long, Key() As Byte
    
    Dim MinKeyLength As TWOFISHKEYLENGTH
    '
    ' Always 256 Key for really strong encryption
    MinKeyLength = TWOFISH_256
    
    KeyLength = Len(New_Value) * 8
    Key() = StrConv(New_Value, vbFromUnicode)
    If (KeyLength < MinKeyLength) Then
        ReDim Preserve Key(MinKeyLength \ 8 - 1)
        KeyLength = MinKeyLength
    End If
    If (KeyLength > 192) Then
        ReDim Preserve Key(31)
        KeyLength = 256
    ElseIf (KeyLength > 128) Then
        ReDim Preserve Key(23)
        KeyLength = 192
    ElseIf (KeyLength > 64) Then
        ReDim Preserve Key(15)
        KeyLength = 128
    ElseIf (KeyLength > 32) Then
        ReDim Preserve Key(7)
        KeyLength = 64
    Else
        ReDim Preserve Key(3)
        KeyLength = 32
    End If
    Call KeyCreate(Key, KeyLength \ 8)
    
    Exit Property
errhandler:
     
    ' ---------------------------------------------------------------------------
    ' Raise friendly error to the handler
    ' ---------------------------------------------------------------------------
    Call Err.Raise(Err.Number, Err.Source, Err.Description)
End Property

Public Function EncryptByte(ByteArray() As Byte, Optional Key As String)
    Dim Offset As Long, OrigLen As Long, CipherLen As Long, CurrPercent As Long, NextPercent As Long, DWord(0 To 5) As Long, CipherWord(0 To 3) As Long
    
  
    On Error GoTo errhandler
    
    If (Len(Key) > 0) Then IAlgorithm_Key = Key
    OrigLen = UBound(ByteArray) + 1
    CipherLen = OrigLen + 12
    If (CipherLen Mod 16 <> 0) Then CipherLen = CipherLen + 16 - (CipherLen Mod 16)
    ReDim Preserve ByteArray(CipherLen - 1)
    Call CopyMem(ByteArray(12), ByteArray(0), OrigLen)
    Call CopyMem(ByteArray(8), OrigLen, 4)
    Call Randomize
    Call CopyMem(ByteArray(0), CLng(2147483647 * Rnd), 4)
    Call CopyMem(ByteArray(4), CLng(2147483647 * Rnd), 4)
    For Offset = 0 To (CipherLen - 1) Step 16
          Call CopyMem(DWord(0), ByteArray(Offset), 16)
          DWord(0) = DWord(0) Xor CipherWord(0)
          DWord(1) = DWord(1) Xor CipherWord(1)
          DWord(2) = DWord(2) Xor CipherWord(2)
          DWord(3) = DWord(3) Xor CipherWord(3)
          Call EncryptBlock(DWord())
          Call CopyMem(ByteArray(Offset), DWord(0), 16)
          Call CopyMem(CipherWord(0), DWord(0), 16)
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
   
    Dim Offset As Long, OrigLen As Long, CipherLen As Long, CurrPercent As Long, NextPercent As Long, DWord(0 To 5) As Long, CipherWord(0 To 3) As Long
    
    On Error GoTo errhandler

    
    If (Len(Key) > 0) Then IAlgorithm_Key = Key
    CipherLen = UBound(ByteArray) + 1
    For Offset = 0 To (CipherLen - 1) Step 16
        Call CopyMem(DWord(2), ByteArray(Offset), 16)
        Call DecryptBlock(DWord())
        DWord(0) = DWord(0) Xor CipherWord(0)
        DWord(1) = DWord(1) Xor CipherWord(1)
        DWord(2) = DWord(2) Xor CipherWord(2)
        DWord(3) = DWord(3) Xor CipherWord(3)
        Call CopyMem(CipherWord(0), ByteArray(Offset), 16)
        Call CopyMem(ByteArray(Offset), DWord(0), 16)
        If (Offset >= NextPercent) Then
            CurrPercent = Int((Offset / CipherLen) * 100)
            NextPercent = (CipherLen * ((CurrPercent + 1) / 100)) + 1
            RaiseEvent Process(CurrPercent)
        End If
        DoEvents
    Next
    Call CopyMem(OrigLen, ByteArray(8), 4)
    If (CipherLen - OrigLen > 27) Or (CipherLen - OrigLen < 12) Then Call Err.Raise(vbObjectError, , "Incorrect size descriptor in Twofish decryption")
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
    If FileExist(OutFile) Then Kill OutFile
    FileO = FreeFile
    Open OutFile For Binary As #FileO
        Put #FileO, , Buffer()
    Close #FileO
    IAlgorithm_EncryptFile = True
    
    m_status = STAT_READY
    RaiseEvent statuschanged(m_status)
    
    Exit Function
errhandler:
    ' ---------------------------------------------------------------------------
    ' Raise friendly error to the handler
    ' ---------------------------------------------------------------------------
    m_status = STAT_ERROR
    RaiseEvent statuschanged(m_status)
    Call Err.Raise(Err.Number, Err.Source, Err.Description)
End Function

Private Function IAlgorithm_DecryptFile(InFile As String, OutFile As String, Overwrite As Boolean, Optional Key As String) As Boolean

    On Error GoTo errhandler
    m_status = STAT_BUSY
    RaiseEvent statuschanged(m_status)
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
    
    If FileExist(OutFile) = True Then Kill OutFile
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
    
    On Error GoTo errhandler
    m_status = STAT_BUSY
    RaiseEvent statuschanged(m_status)
    
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
    On Error GoTo errhandler
    m_status = STAT_BUSY
    RaiseEvent statuschanged(m_status)
    If IsTextInHex = True Then Text = DeHex(Text)
    Dim ByteArray() As Byte
    ByteArray() = StrConv(Text, vbFromUnicode)
    Call DecryptByte(ByteArray(), Key)
    IAlgorithm_DecryptString = StrConv(ByteArray(), vbUnicode)
       
    Exit Function
errhandler:
    
    ' ---------------------------------------------------------------------------
    ' Raise friendly error to the handler
    ' ---------------------------------------------------------------------------
    Call Err.Raise(Err.Number, Err.Source, Err.Description)
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
Private Function DeHex(Data As String) As String
    Dim iCount As Double
    Reset
    For iCount = 1 To Len(Data) Step 2
        Append Chr$(Val("&H" & Mid$(Data, iCount, 2)))
    Next
    DeHex = GData
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
Private Static Sub GetWord(LongValue As Long, CryptBuffer() As Byte, Offset As Long)
    Dim bb(0 To 3) As Byte
    bb(3) = CryptBuffer(Offset)
    bb(2) = CryptBuffer(Offset + 1)
    bb(1) = CryptBuffer(Offset + 2)
    bb(0) = CryptBuffer(Offset + 3)
    Call CopyMem(LongValue, bb(0), 4)
End Sub
Private Static Sub PutWord(LongValue As Long, CryptBuffer() As Byte, Offset As Long)
    Dim bb(0 To 3) As Byte
    Call CopyMem(bb(0), LongValue, 4)
    CryptBuffer(Offset) = bb(3)
    CryptBuffer(Offset + 1) = bb(2)
    CryptBuffer(Offset + 2) = bb(1)
    CryptBuffer(Offset + 3) = bb(0)
End Sub
Private Static Function UnsignedAdd(ByVal Data1 As Long, Data2 As Long) As Long
    Dim X1(0 To 3) As Byte, X2(0 To 3) As Byte, xx(0 To 3) As Byte, Rest As Long, Value As Long, a As Long
    Call CopyMem(X1(0), Data1, 4)
    Call CopyMem(X2(0), Data2, 4)
    Rest = 0
    For a = 0 To 3
        Value = CLng(X1(a)) + CLng(X2(a)) + Rest
        xx(a) = Value And 255
        Rest = Value \ 256
    Next
    Call CopyMem(UnsignedAdd, xx(0), 4)
End Function
Private Function UnsignedDel(Data1 As Long, Data2 As Long) As Long
    Dim X1(0 To 3) As Byte, X2(0 To 3) As Byte, xx(0 To 3) As Byte, Rest As Long, Value As Long, a As Long
    Call CopyMem(X1(0), Data1, 4)
    Call CopyMem(X2(0), Data2, 4)
    Call CopyMem(xx(0), UnsignedDel, 4)
    For a = 0 To 3
        Value = CLng(X1(a)) - CLng(X2(a)) - Rest
        If (Value < 0) Then
            Value = Value + 256
            Rest = 1
        Else
            Rest = 0
        End If
        xx(a) = Value
    Next
    Call CopyMem(UnsignedDel, xx(0), 4)
End Function

Private Static Function LFSR1(ByRef x As Long) As Long
    LFSR1 = lBSR(x, 1) Xor ((x And 1) * GF256_FDBK_2)
End Function
Private Static Function LFSR2(ByRef x As Long) As Long
    LFSR2 = lBSR(x, 2) Xor ((x And &H2) / &H2 * GF256_FDBK_2) Xor ((x And &H1) * GF256_FDBK_4)
End Function
Private Static Function RS_Rem(x As Long) As Long
    Dim B As Long, g2 As Long, g3 As Long
    B = (lBSRU(x, 24) And &HFF)
    g2 = ((lBSL(B, 1) Xor (B And &H80) / &H80 * &H14D) And &HFF)
    g3 = (lBSRU(B, 1) Xor ((B And &H1) * lBSRU(&H14D, 1)) Xor g2)
    RS_Rem = lBSL(x, 8) Xor lBSL(g3, 24) Xor lBSL(g2, 16) Xor lBSL(g3, 8) Xor B
End Function
Private Static Function F32(k64Cnt As Long, x As Long, k32() As Long) As Long
    Dim xb(0 To 3) As Byte, Key(0 To 3, 0 To 3) As Byte
    Call CopyMem(xb(0), x, 4)
    Call CopyMem(Key(0, 0), k32(0), 16)
    If ((k64Cnt And 3) = 1) Then
        F32 = MDS(0, p(0, xb(0)) Xor Key(0, 0)) Xor MDS(1, p(0, xb(1)) Xor Key(1, 0)) Xor MDS(2, p(1, xb(2)) Xor Key(2, 0)) Xor MDS(3, p(1, xb(3)) Xor Key(3, 0))
    Else
        If ((k64Cnt And 3) = 0) Then
              xb(0) = p(1, xb(0)) Xor Key(0, 3)
              xb(1) = p(0, xb(1)) Xor Key(1, 3)
              xb(2) = p(0, xb(2)) Xor Key(2, 3)
              xb(3) = p(1, xb(3)) Xor Key(3, 3)
        End If
        If ((k64Cnt And 3) = 3) Or ((k64Cnt And 3) = 0) Then
              xb(0) = p(1, xb(0)) Xor Key(0, 2)
              xb(1) = p(1, xb(1)) Xor Key(1, 2)
              xb(2) = p(0, xb(2)) Xor Key(2, 2)
              xb(3) = p(0, xb(3)) Xor Key(3, 2)
        End If
        F32 = MDS(0, p(0, p(0, xb(0)) Xor Key(0, 1)) Xor Key(0, 0)) Xor MDS(1, p(0, p(1, xb(1)) Xor Key(1, 1)) Xor Key(1, 0)) Xor MDS(2, p(1, p(0, xb(2)) Xor Key(2, 1)) Xor Key(2, 0)) Xor MDS(3, p(1, p(1, xb(3)) Xor Key(3, 1)) Xor Key(3, 0))
    End If
End Function
Private Static Function Fe32(x As Long, R As Long) As Long
    Dim xb(0 To 3) As Byte
    Call CopyMem(xb(0), x, 4)
    Fe32 = sBox(2 * xb(R Mod 4)) Xor sBox(2 * xb((R + 1) Mod 4) + 1) Xor sBox(&H200 + 2 * xb((R + 2) Mod 4)) Xor sBox(&H200 + 2 * xb((R + 3) Mod 4) + 1)
End Function
Private Static Sub KeyCreate(K() As Byte, KeyLength As Long)
    Dim i As Long, lA As Long, lB As Long, B(3) As Byte, k64Cnt As Long, k32e(3) As Long, k32o(3) As Long, subkeyCnt As Long, sBoxKey(3) As Long, Key(0 To 3, 0 To 3) As Byte
    k64Cnt = KeyLength \ 8
    subkeyCnt = ROUND_SUBKEYS + 2 * ROUNDS
    For i = 0 To IIf(KeyLength < 32, KeyLength \ 8 - 1, 3)
        Call CopyMem(k32e(i), K(i * 8), 4)
        Call CopyMem(k32o(i), K(i * 8 + 4), 4)
        sBoxKey(KeyLength \ 8 - 1 - i) = RS_Rem(RS_Rem(RS_Rem(RS_Rem(RS_Rem(RS_Rem(RS_Rem(RS_Rem(k32o(i))))) Xor k32e(i)))))
    Next
    ReDim sKey(subkeyCnt)
    For i = 0 To ((subkeyCnt / 2) - 1)
        lA = F32(k64Cnt, i * &H2020202, k32e)
        lB = F32(k64Cnt, i * &H2020202 + &H1010101, k32o)
        lB = lBSL(lB, 8) Or lBSRU(lB, 24)
        If (m_RunningCompiled) Then lA = lA + lB Else lA = UnsignedAdd(lA, lB)
        sKey(2 * i) = lA
        If (m_RunningCompiled) Then lA = lA + lB Else lA = UnsignedAdd(lA, lB)
        sKey(2 * i + 1) = lBSL(lA, 9) Or lBSRU(lA, 32 - 9)
    Next
    Call CopyMem(Key(0, 0), sBoxKey(0), 16)
    For i = 0 To 255
        If ((k64Cnt And 3) = 1) Then
            sBox(2 * i) = MDS(0, p(0, i) Xor Key(0, 0))
            sBox(2 * i + 1) = MDS(1, p(0, i) Xor Key(1, 0))
            sBox(&H200 + 2 * i) = MDS(2, p(1, i) Xor Key(2, 0))
            sBox(&H200 + 2 * i + 1) = MDS(3, p(1, i) Xor Key(3, 0))
        Else
            B(0) = i
            B(1) = i
            B(2) = i
            B(3) = i
            If ((k64Cnt And 3) = 0) Then
                  B(0) = p(1, B(0)) Xor Key(0, 3)
                  B(1) = p(0, B(1)) Xor Key(1, 3)
                  B(2) = p(0, B(2)) Xor Key(2, 3)
                  B(3) = p(1, B(3)) Xor Key(3, 3)
            End If
            If ((k64Cnt And 3) = 3) Or ((k64Cnt And 3) = 0) Then
                B(0) = p(1, B(0)) Xor Key(0, 2)
                B(1) = p(1, B(1)) Xor Key(1, 2)
                B(2) = p(0, B(2)) Xor Key(2, 2)
                B(3) = p(0, B(3)) Xor Key(3, 2)
            End If
            sBox(2 * i) = MDS(0, p(0, p(0, B(0)) Xor Key(0, 1)) Xor Key(0, 0))
            sBox(2 * i + 1) = MDS(1, p(0, p(1, B(1)) Xor Key(1, 1)) Xor Key(1, 0))
            sBox(&H200 + 2 * i) = MDS(2, p(1, p(0, B(2)) Xor Key(2, 1)) Xor Key(2, 0))
            sBox(&H200 + 2 * i + 1) = MDS(3, p(1, p(1, B(3)) Xor Key(3, 1)) Xor Key(3, 0))
        End If
    Next
End Sub
Private Function lBSL(ByRef lInput As Long, ByRef bShiftBits As Byte) As Long
    lBSL = (lInput And (2 ^ (31 - bShiftBits) - 1)) * 2 ^ bShiftBits
    If (lInput And 2 ^ (31 - bShiftBits)) = 2 ^ (31 - bShiftBits) Then lBSL = (lBSL Or &H80000000)
End Function
Private Function lBSR(ByRef lInput As Long, ByRef bShiftBits As Byte) As Long
    If (bShiftBits = 31) Then
        If (lInput < 0) Then lBSR = &HFFFFFFFF Else lBSR = 0
    Else
        lBSR = (lInput And Not (2 ^ bShiftBits - 1)) \ 2 ^ bShiftBits
    End If
End Function
Private Function lBSRU(lInput As Long, bShiftBits As Byte) As Long
    If (bShiftBits = 31) Then
        lBSRU = -(lInput < 0)
    Else
        lBSRU = (((lInput And Not (2 ^ bShiftBits - 1)) \ 2 ^ bShiftBits) And Not (&H80000000 + (2 ^ bShiftBits - 2) * 2 ^ (31 - bShiftBits)))
    End If
End Function
Private Static Sub EncryptBlock(DWord() As Long)
    Dim t0 As Long, T1 As Long, K As Long, R As Long
    DWord(0) = DWord(0) Xor sKey(INPUT_WHITEN)
    DWord(1) = DWord(1) Xor sKey(INPUT_WHITEN + 1)
    DWord(2) = DWord(2) Xor sKey(INPUT_WHITEN + 2)
    DWord(3) = DWord(3) Xor sKey(INPUT_WHITEN + 3)
    K = ROUND_SUBKEYS
    For R = 0 To (ROUNDS - 1) Step 2
        If (m_RunningCompiled) Then
            t0 = Fe32(DWord(0), 0)
            T1 = Fe32(DWord(1), 3)
            t0 = t0 + T1
            DWord(2) = Rot1(DWord(2) Xor (t0 + sKey(K)))
            K = K + 1
            DWord(3) = Rot31(DWord(3)) Xor (t0 + T1 + sKey(K))
            K = K + 1
            t0 = Fe32(DWord(2), 0)
            T1 = Fe32(DWord(3), 3)
            t0 = t0 + T1
            DWord(0) = Rot1(DWord(0) Xor (t0 + sKey(K)))
            K = K + 1
            DWord(1) = Rot31(DWord(1)) Xor (t0 + T1 + sKey(K))
            K = K + 1
        Else
            t0 = Fe32(DWord(0), 0)
            T1 = Fe32(DWord(1), 3)
            t0 = UnsignedAdd(t0, T1)
            DWord(2) = Rot1(DWord(2) Xor (UnsignedAdd(t0, sKey(K))))
            K = K + 1
            DWord(3) = Rot31(DWord(3)) Xor (UnsignedAdd(UnsignedAdd(t0, T1), sKey(K)))
            K = K + 1
            t0 = Fe32(DWord(2), 0)
            T1 = Fe32(DWord(3), 3)
            t0 = UnsignedAdd(t0, T1)
            DWord(0) = Rot1(DWord(0) Xor (UnsignedAdd(t0, sKey(K))))
            K = K + 1
            DWord(1) = Rot31(DWord(1)) Xor (UnsignedAdd(UnsignedAdd(t0, T1), sKey(K)))
            K = K + 1
        End If
    Next
    DWord(2) = DWord(2) Xor sKey(OUTPUT_WHITEN)
    DWord(3) = DWord(3) Xor sKey(OUTPUT_WHITEN + 1)
    DWord(4) = DWord(0) Xor sKey(OUTPUT_WHITEN + 2)
    DWord(5) = DWord(1) Xor sKey(OUTPUT_WHITEN + 3)
    Call CopyMem(DWord(0), DWord(2), 16)
End Sub
Private Sub DecryptBlock(DWord() As Long)
    Dim K As Long, R As Long, t0 As Long, T1 As Long
    DWord(2) = DWord(2) Xor sKey(OUTPUT_WHITEN)
    DWord(3) = DWord(3) Xor sKey(OUTPUT_WHITEN + 1)
    DWord(0) = DWord(4) Xor sKey(OUTPUT_WHITEN + 2)
    DWord(1) = DWord(5) Xor sKey(OUTPUT_WHITEN + 3)
    K = ROUND_SUBKEYS + 2 * ROUNDS - 1
    For R = 0 To ROUNDS - 1 Step 2
        If (m_RunningCompiled) Then
            t0 = Fe32(DWord(2), 0)
            T1 = Fe32(DWord(3), 3)
            t0 = t0 + T1
            DWord(1) = Rot1(DWord(1) Xor (t0 + T1 + sKey(K)))
            K = K - 1
            DWord(0) = Rot31(DWord(0)) Xor (t0 + sKey(K))
            K = K - 1
            t0 = Fe32(DWord(0), 0)
            T1 = Fe32(DWord(1), 3)
            t0 = t0 + T1
            DWord(3) = Rot1(DWord(3) Xor (t0 + T1 + sKey(K)))
            K = K - 1
            DWord(2) = Rot31(DWord(2)) Xor (t0 + sKey(K))
            K = K - 1
        Else
            t0 = Fe32(DWord(2), 0)
            T1 = Fe32(DWord(3), 3)
            t0 = UnsignedAdd(t0, T1)
            DWord(1) = Rot1(DWord(1) Xor (UnsignedAdd(UnsignedAdd(t0, T1), sKey(K))))
            K = K - 1
            DWord(0) = Rot31(DWord(0)) Xor (UnsignedAdd(t0, sKey(K)))
            K = K - 1
            t0 = Fe32(DWord(0), 0)
            T1 = Fe32(DWord(1), 3)
            t0 = UnsignedAdd(t0, T1)
            DWord(3) = Rot1(DWord(3) Xor (UnsignedAdd(UnsignedAdd(t0, T1), sKey(K))))
            K = K - 1
            DWord(2) = Rot31(DWord(2)) Xor (UnsignedAdd(t0, sKey(K)))
            K = K - 1
        End If
    Next
    DWord(0) = DWord(0) Xor sKey(INPUT_WHITEN)
    DWord(1) = DWord(1) Xor sKey(INPUT_WHITEN + 1)
    DWord(2) = DWord(2) Xor sKey(INPUT_WHITEN + 2)
    DWord(3) = DWord(3) Xor sKey(INPUT_WHITEN + 3)
End Sub
Private Static Function Rot1(Value As Long) As Long
    Dim temp As Byte, x(0 To 3) As Byte
    Call CopyMem(x(0), Value, 4)
    temp = x(0)
    x(0) = (x(0) \ 2) Or ((x(1) And 1) * 128)
    x(1) = (x(1) \ 2) Or ((x(2) And 1) * 128)
    x(2) = (x(2) \ 2) Or ((x(3) And 1) * 128)
    x(3) = (x(3) \ 2) Or ((temp And 1) * 128)
    Call CopyMem(Rot1, x(0), 4)
End Function
Private Static Function Rot31(Value As Long) As Long
    Dim temp As Byte, x(0 To 3) As Byte
    Call CopyMem(x(0), Value, 4)
    temp = x(3)
    x(3) = ((x(3) And 127) * 2) Or -CBool(x(2) And 128)
    x(2) = ((x(2) And 127) * 2) Or -CBool(x(1) And 128)
    x(1) = ((x(1) And 127) * 2) Or -CBool(x(0) And 128)
    x(0) = ((x(0) And 127) * 2) Or -CBool(temp And 128)
    Call CopyMem(Rot31, x(0), 4)
End Function
Private Sub Class_Initialize()
Dim i As Long, j As Long, m1(0 To 1) As Long, mX(0 To 1) As Long, mY(0 To 1) As Long
  
On Local Error Resume Next
m_RunningCompiled = ((2147483647 + 1) < 0)

  p(0, 0) = &HA9
  p(0, 1) = &H67
  p(0, 2) = &HB3
  p(0, 3) = &HE8
  p(0, 4) = &H4
  p(0, 5) = &HFD
  p(0, 6) = &HA3
  p(0, 7) = &H76
  p(0, 8) = &H9A
  p(0, 9) = &H92
  p(0, 10) = &H80
  p(0, 11) = &H78
  p(0, 12) = &HE4
  p(0, 13) = &HDD
  p(0, 14) = &HD1
  p(0, 15) = &H38
  p(0, 16) = &HD
  p(0, 17) = &HC6
  p(0, 18) = &H35
  p(0, 19) = &H98
  p(0, 20) = &H18
  p(0, 21) = &HF7
  p(0, 22) = &HEC
  p(0, 23) = &H6C
  p(0, 24) = &H43
  p(0, 25) = &H75
  p(0, 26) = &H37
  p(0, 27) = &H26
  p(0, 28) = &HFA
  p(0, 29) = &H13
  p(0, 30) = &H94
  p(0, 31) = &H48
  p(0, 32) = &HF2
  p(0, 33) = &HD0
  p(0, 34) = &H8B
  p(0, 35) = &H30
  p(0, 36) = &H84
  p(0, 37) = &H54
  p(0, 38) = &HDF
  p(0, 39) = &H23
  p(0, 40) = &H19
  p(0, 41) = &H5B
  p(0, 42) = &H3D
  p(0, 43) = &H59
  p(0, 44) = &HF3
  p(0, 45) = &HAE
  p(0, 46) = &HA2
  p(0, 47) = &H82
  p(0, 48) = &H63
  p(0, 49) = &H1
  p(0, 50) = &H83
  p(0, 51) = &H2E
  p(0, 52) = &HD9
  p(0, 53) = &H51
  p(0, 54) = &H9B
  p(0, 55) = &H7C
  p(0, 56) = &HA6
  p(0, 57) = &HEB
  p(0, 58) = &HA5
  p(0, 59) = &HBE
  p(0, 60) = &H16
  p(0, 61) = &HC
  p(0, 62) = &HE3
  p(0, 63) = &H61
  p(0, 64) = &HC0
  p(0, 65) = &H8C
  p(0, 66) = &H3A
  p(0, 67) = &HF5
  p(0, 68) = &H73
  p(0, 69) = &H2C
  p(0, 70) = &H25
  p(0, 71) = &HB
  p(0, 72) = &HBB
  p(0, 73) = &H4E
  p(0, 74) = &H89
  p(0, 75) = &H6B
  p(0, 76) = &H53
  p(0, 77) = &H6A
  p(0, 78) = &HB4
  p(0, 79) = &HF1
  p(0, 80) = &HE1
  p(0, 81) = &HE6
  p(0, 82) = &HBD
  p(0, 83) = &H45
  p(0, 84) = &HE2
  p(0, 85) = &HF4
  p(0, 86) = &HB6
  p(0, 87) = &H66
  p(0, 88) = &HCC
  p(0, 89) = &H95
  p(0, 90) = &H3
  p(0, 91) = &H56
  p(0, 92) = &HD4
  p(0, 93) = &H1C
  p(0, 94) = &H1E
  p(0, 95) = &HD7
  p(0, 96) = &HFB
  p(0, 97) = &HC3
  p(0, 98) = &H8E
  p(0, 99) = &HB5
  p(0, 100) = &HE9
  p(0, 101) = &HCF
  p(0, 102) = &HBF
  p(0, 103) = &HBA
  p(0, 104) = &HEA
  p(0, 105) = &H77
  p(0, 106) = &H39
  p(0, 107) = &HAF
  p(0, 108) = &H33
  p(0, 109) = &HC9
  p(0, 110) = &H62
  p(0, 111) = &H71
  p(0, 112) = &H81
  p(0, 113) = &H79
  p(0, 114) = &H9
  p(0, 115) = &HAD
  p(0, 116) = &H24
  p(0, 117) = &HCD
  p(0, 118) = &HF9
  p(0, 119) = &HD8
  p(0, 120) = &HE5
  p(0, 121) = &HC5
  p(0, 122) = &HB9
  p(0, 123) = &H4D
  p(0, 124) = &H44
  p(0, 125) = &H8
  p(0, 126) = &H86
  p(0, 127) = &HE7
  p(0, 128) = &HA1
  p(0, 129) = &H1D
  p(0, 130) = &HAA
  p(0, 131) = &HED
  p(0, 132) = &H6
  p(0, 133) = &H70
  p(0, 134) = &HB2
  p(0, 135) = &HD2
  p(0, 136) = &H41
  p(0, 137) = &H7B
  p(0, 138) = &HA0
  p(0, 139) = &H11
  p(0, 140) = &H31
  p(0, 141) = &HC2
  p(0, 142) = &H27
  p(0, 143) = &H90
  p(0, 144) = &H20
  p(0, 145) = &HF6
  p(0, 146) = &H60
  p(0, 147) = &HFF
  p(0, 148) = &H96
  p(0, 149) = &H5C
  p(0, 150) = &HB1
  p(0, 151) = &HAB
  p(0, 152) = &H9E
  p(0, 153) = &H9C
  p(0, 154) = &H52
  p(0, 155) = &H1B
  p(0, 156) = &H5F
  p(0, 157) = &H93
  p(0, 158) = &HA
  p(0, 159) = &HEF
  p(0, 160) = &H91
  p(0, 161) = &H85
  p(0, 162) = &H49
  p(0, 163) = &HEE
  p(0, 164) = &H2D
  p(0, 165) = &H4F
  p(0, 166) = &H8F
  p(0, 167) = &H3B
  p(0, 168) = &H47
  p(0, 169) = &H87
  p(0, 170) = &H6D
  p(0, 171) = &H46
  p(0, 172) = &HD6
  p(0, 173) = &H3E
  p(0, 174) = &H69
  p(0, 175) = &H64
  p(0, 176) = &H2A
  p(0, 177) = &HCE
  p(0, 178) = &HCB
  p(0, 179) = &H2F
  p(0, 180) = &HFC
  p(0, 181) = &H97
  p(0, 182) = &H5
  p(0, 183) = &H7A
  p(0, 184) = &HAC
  p(0, 185) = &H7F
  p(0, 186) = &HD5
  p(0, 187) = &H1A
  p(0, 188) = &H4B
  p(0, 189) = &HE
  p(0, 190) = &HA7
  p(0, 191) = &H5A
  p(0, 192) = &H28
  p(0, 193) = &H14
  p(0, 194) = &H3F
  p(0, 195) = &H29
  p(0, 196) = &H88
  p(0, 197) = &H3C
  p(0, 198) = &H4C
  p(0, 199) = &H2
  p(0, 200) = &HB8
  p(0, 201) = &HDA
  p(0, 202) = &HB0
  p(0, 203) = &H17
  p(0, 204) = &H55
  p(0, 205) = &H1F
  p(0, 206) = &H8A
  p(0, 207) = &H7D
  p(0, 208) = &H57
  p(0, 209) = &HC7
  p(0, 210) = &H8D
  p(0, 211) = &H74
  p(0, 212) = &HB7
  p(0, 213) = &HC4
  p(0, 214) = &H9F
  p(0, 215) = &H72
  p(0, 216) = &H7E
  p(0, 217) = &H15
  p(0, 218) = &H22
  p(0, 219) = &H12
  p(0, 220) = &H58
  p(0, 221) = &H7
  p(0, 222) = &H99
  p(0, 223) = &H34
  p(0, 224) = &H6E
  p(0, 225) = &H50
  p(0, 226) = &HDE
  p(0, 227) = &H68
  p(0, 228) = &H65
  p(0, 229) = &HBC
  p(0, 230) = &HDB
  p(0, 231) = &HF8
  p(0, 232) = &HC8
  p(0, 233) = &HA8
  p(0, 234) = &H2B
  p(0, 235) = &H40
  p(0, 236) = &HDC
  p(0, 237) = &HFE
  p(0, 238) = &H32
  p(0, 239) = &HA4
  p(0, 240) = &HCA
  p(0, 241) = &H10
  p(0, 242) = &H21
  p(0, 243) = &HF0
  p(0, 244) = &HD3
  p(0, 245) = &H5D
  p(0, 246) = &HF
  p(0, 247) = &H0
  p(0, 248) = &H6F
  p(0, 249) = &H9D
  p(0, 250) = &H36
  p(0, 251) = &H42
  p(0, 252) = &H4A
  p(0, 253) = &H5E
  p(0, 254) = &HC1
  p(0, 255) = &HE0
  p(1, 0) = &H75
  p(1, 1) = &HF3
  p(1, 2) = &HC6
  p(1, 3) = &HF4
  p(1, 4) = &HDB
  p(1, 5) = &H7B
  p(1, 6) = &HFB
  p(1, 7) = &HC8
  p(1, 8) = &H4A
  p(1, 9) = &HD3
  p(1, 10) = &HE6
  p(1, 11) = &H6B
  p(1, 12) = &H45
  p(1, 13) = &H7D
  p(1, 14) = &HE8
  p(1, 15) = &H4B
  p(1, 16) = &HD6
  p(1, 17) = &H32
  p(1, 18) = &HD8
  p(1, 19) = &HFD
  p(1, 20) = &H37
  p(1, 21) = &H71
  p(1, 22) = &HF1
  p(1, 23) = &HE1
  p(1, 24) = &H30
  p(1, 25) = &HF
  p(1, 26) = &HF8
  p(1, 27) = &H1B
  p(1, 28) = &H87
  p(1, 29) = &HFA
  p(1, 30) = &H6
  p(1, 31) = &H3F
  p(1, 32) = &H5E
  p(1, 33) = &HBA
  p(1, 34) = &HAE
  p(1, 35) = &H5B
  p(1, 36) = &H8A
  p(1, 37) = &H0
  p(1, 38) = &HBC
  p(1, 39) = &H9D
  p(1, 40) = &H6D
  p(1, 41) = &HC1
  p(1, 42) = &HB1
  p(1, 43) = &HE
  p(1, 44) = &H80
  p(1, 45) = &H5D
  p(1, 46) = &HD2
  p(1, 47) = &HD5
  p(1, 48) = &HA0
  p(1, 49) = &H84
  p(1, 50) = &H7
  p(1, 51) = &H14
  p(1, 52) = &HB5
  p(1, 53) = &H90
  p(1, 54) = &H2C
  p(1, 55) = &HA3
  p(1, 56) = &HB2
  p(1, 57) = &H73
  p(1, 58) = &H4C
  p(1, 59) = &H54
  p(1, 60) = &H92
  p(1, 61) = &H74
  p(1, 62) = &H36
  p(1, 63) = &H51
  p(1, 64) = &H38
  p(1, 65) = &HB0
  p(1, 66) = &HBD
  p(1, 67) = &H5A
  p(1, 68) = &HFC
  p(1, 69) = &H60
  p(1, 70) = &H62
  p(1, 71) = &H96
  p(1, 72) = &H6C
  p(1, 73) = &H42
  p(1, 74) = &HF7
  p(1, 75) = &H10
  p(1, 76) = &H7C
  p(1, 77) = &H28
  p(1, 78) = &H27
  p(1, 79) = &H8C
  p(1, 80) = &H13
  p(1, 81) = &H95
  p(1, 82) = &H9C
  p(1, 83) = &HC7
  p(1, 84) = &H24
  p(1, 85) = &H46
  p(1, 86) = &H3B
  p(1, 87) = &H70
  p(1, 88) = &HCA
  p(1, 89) = &HE3
  p(1, 90) = &H85
  p(1, 91) = &HCB
  p(1, 92) = &H11
  p(1, 93) = &HD0
  p(1, 94) = &H93
  p(1, 95) = &HB8
  p(1, 96) = &HA6
  p(1, 97) = &H83
  p(1, 98) = &H20
  p(1, 99) = &HFF
  p(1, 100) = &H9F
  p(1, 101) = &H77
  p(1, 102) = &HC3
  p(1, 103) = &HCC
  p(1, 104) = &H3
  p(1, 105) = &H6F
  p(1, 106) = &H8
  p(1, 107) = &HBF
  p(1, 108) = &H40
  p(1, 109) = &HE7
  p(1, 110) = &H2B
  p(1, 111) = &HE2
  p(1, 112) = &H79
  p(1, 113) = &HC
  p(1, 114) = &HAA
  p(1, 115) = &H82
  p(1, 116) = &H41
  p(1, 117) = &H3A
  p(1, 118) = &HEA
  p(1, 119) = &HB9
  p(1, 120) = &HE4
  p(1, 121) = &H9A
  p(1, 122) = &HA4
  p(1, 123) = &H97
  p(1, 124) = &H7E
  p(1, 125) = &HDA
  p(1, 126) = &H7A
  p(1, 127) = &H17
  p(1, 128) = &H66
  p(1, 129) = &H94
  p(1, 130) = &HA1
  p(1, 131) = &H1D
  p(1, 132) = &H3D
  p(1, 133) = &HF0
  p(1, 134) = &HDE
  p(1, 135) = &HB3
  p(1, 136) = &HB
  p(1, 137) = &H72
  p(1, 138) = &HA7
  p(1, 139) = &H1C
  p(1, 140) = &HEF
  p(1, 141) = &HD1
  p(1, 142) = &H53
  p(1, 143) = &H3E
  p(1, 144) = &H8F
  p(1, 145) = &H33
  p(1, 146) = &H26
  p(1, 147) = &H5F
  p(1, 148) = &HEC
  p(1, 149) = &H76
  p(1, 150) = &H2A
  p(1, 151) = &H49
  p(1, 152) = &H81
  p(1, 153) = &H88
  p(1, 154) = &HEE
  p(1, 155) = &H21
  p(1, 156) = &HC4
  p(1, 157) = &H1A
  p(1, 158) = &HEB
  p(1, 159) = &HD9
  p(1, 160) = &HC5
  p(1, 161) = &H39
  p(1, 162) = &H99
  p(1, 163) = &HCD
  p(1, 164) = &HAD
  p(1, 165) = &H31
  p(1, 166) = &H8B
  p(1, 167) = &H1
  p(1, 168) = &H18
  p(1, 169) = &H23
  p(1, 170) = &HDD
  p(1, 171) = &H1F
  p(1, 172) = &H4E
  p(1, 173) = &H2D
  p(1, 174) = &HF9
  p(1, 175) = &H48
  p(1, 176) = &H4F
  p(1, 177) = &HF2
  p(1, 178) = &H65
  p(1, 179) = &H8E
  p(1, 180) = &H78
  p(1, 181) = &H5C
  p(1, 182) = &H58
  p(1, 183) = &H19
  p(1, 184) = &H8D
  p(1, 185) = &HE5
  p(1, 186) = &H98
  p(1, 187) = &H57
  p(1, 188) = &H67
  p(1, 189) = &H7F
  p(1, 190) = &H5
  p(1, 191) = &H64
  p(1, 192) = &HAF
  p(1, 193) = &H63
  p(1, 194) = &HB6
  p(1, 195) = &HFE
  p(1, 196) = &HF5
  p(1, 197) = &HB7
  p(1, 198) = &H3C
  p(1, 199) = &HA5
  p(1, 200) = &HCE
  p(1, 201) = &HE9
  p(1, 202) = &H68
  p(1, 203) = &H44
  p(1, 204) = &HE0
  p(1, 205) = &H4D
  p(1, 206) = &H43
  p(1, 207) = &H69
  p(1, 208) = &H29
  p(1, 209) = &H2E
  p(1, 210) = &HAC
  p(1, 211) = &H15
  p(1, 212) = &H59
  p(1, 213) = &HA8
  p(1, 214) = &HA
  p(1, 215) = &H9E
  p(1, 216) = &H6E
  p(1, 217) = &H47
  p(1, 218) = &HDF
  p(1, 219) = &H34
  p(1, 220) = &H35
  p(1, 221) = &H6A
  p(1, 222) = &HCF
  p(1, 223) = &HDC
  p(1, 224) = &H22
  p(1, 225) = &HC9
  p(1, 226) = &HC0
  p(1, 227) = &H9B
  p(1, 228) = &H89
  p(1, 229) = &HD4
  p(1, 230) = &HED
  p(1, 231) = &HAB
  p(1, 232) = &H12
  p(1, 233) = &HA2
  p(1, 234) = &HD
  p(1, 235) = &H52
  p(1, 236) = &HBB
  p(1, 237) = &H2
  p(1, 238) = &H2F
  p(1, 239) = &HA9
  p(1, 240) = &HD7
  p(1, 241) = &H61
  p(1, 242) = &H1E
  p(1, 243) = &HB4
  p(1, 244) = &H50
  p(1, 245) = &H4
  p(1, 246) = &HF6
  p(1, 247) = &HC2
  p(1, 248) = &H16
  p(1, 249) = &H25
  p(1, 250) = &H86
  p(1, 251) = &H56
  p(1, 252) = &H55
  p(1, 253) = &H9
  p(1, 254) = &HBE
  p(1, 255) = &H91

  For i = 0 To 255
        j = p(0, i)
        m1(0) = j
        mX(0) = j Xor LFSR2(j)
        mY(0) = j Xor LFSR1(j) Xor LFSR2(j)
        j = p(1, i)
        m1(1) = j
        mX(1) = j Xor LFSR2(j)
        mY(1) = j Xor LFSR1(j) Xor LFSR2(j)
        MDS(0, i) = (m1(1) Or lBSL(mX(1), 8) Or lBSL(mY(1), 16) Or lBSL(mY(1), 24))
        MDS(1, i) = (mY(0) Or lBSL(mY(0), 8) Or lBSL(mX(0), 16) Or lBSL(m1(0), 24))
        MDS(2, i) = (mX(1) Or lBSL(mY(1), 8) Or lBSL(m1(1), 16) Or lBSL(mY(1), 24))
        MDS(3, i) = (mX(0) Or lBSL(m1(0), 8) Or lBSL(mY(0), 16) Or lBSL(mX(0), 24))
  Next
  m_status = STAT_READY
End Sub

Private Property Let IAlgorithm_Status(lStatus As Long)

    m_status = lStatus
    RaiseEvent statuschanged(lStatus)
 
End Property

Private Property Get IAlgorithm_Status() As Long
    IAlgorithm_Status = m_status
End Property

Private Function IAlgorithm_ToString() As String
    IAlgorithm_ToString = "Twofish"
End Function
