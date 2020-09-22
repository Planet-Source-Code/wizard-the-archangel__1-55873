VERSION 5.00
Begin VB.UserControl CryptoEngine 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   ClientHeight    =   1095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1215
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   1095
   ScaleWidth      =   1215
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "ActiveX Cryptography Control"
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "CryptoEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum enuCRYPTO_ALGORITHMS
    acuBlowfish
    acuDES
    acuGost
    acuRijndael
    acuSkipjack
    acuTEA
    acuTwofish
End Enum


Public Enum enuHASH_ALGORITHMS
    acuMD5
    acuSHA256
End Enum

Public Enum enuSECURITY_STATUS
    acuHigh
    acuMedium
    acuLow
End Enum

Private m_HashType          As enuHASH_ALGORITHMS
Private m_AlgoType          As enuCRYPTO_ALGORITHMS
Private m_SecStatus         As enuSECURITY_STATUS
Private m_status            As Long
Private m_Engine            As IAlgorithm
Private m_Hash              As IHashAlgorithm
Private m_Key               As String

Event Process(percent As Long)
Event statuschanged(lStatus As Long)
Event Error(Number As Long, Source As String, Description As String)

Private WithEvents m_Blowfish As Blowfish
Attribute m_Blowfish.VB_VarHelpID = -1
Private WithEvents m_DES      As DES
Attribute m_DES.VB_VarHelpID = -1
Private WithEvents m_Gost     As Gost
Attribute m_Gost.VB_VarHelpID = -1
Private WithEvents m_Rijndael As Rijndael
Attribute m_Rijndael.VB_VarHelpID = -1
Private WithEvents m_SkipJack As Skipjack
Attribute m_SkipJack.VB_VarHelpID = -1
Private WithEvents m_Twofish  As Twofish
Attribute m_Twofish.VB_VarHelpID = -1
Private WithEvents m_TEA      As TEA
Attribute m_TEA.VB_VarHelpID = -1

Public Property Get HashAlgorithm() As enuHASH_ALGORITHMS
    HashAlgorithm = m_HashType
End Property

Public Property Let HashAlgorithm(ByVal enuHash As enuHASH_ALGORITHMS)
    If enuHash > acuSHA256 Or enuHash < acuMD5 Then
        Err.Raise 1001, "CryptoEngine:HashAlgorithm", "Invalid property value"
    Else
        m_HashType = enuHash
        SetHash m_HashType
    End If
    
End Property

Public Property Get CryptAlgorithm() As enuCRYPTO_ALGORITHMS
    CryptAlgorithm = m_AlgoType
End Property

Public Property Let CryptAlgorithm(ByVal enuCrypt As enuCRYPTO_ALGORITHMS)
    On Error GoTo errhandler
    If enuCrypt > acuTwofish Or enuCrypt < acuBlowfish Then
        Err.Raise 1001, "CryptoEngine:CryptAlgorithm", "Invalid property value"
    Else
        If SetEngine(m_AlgoType) Then m_AlgoType = enuCrypt
        m_status = STAT_READY
        RaiseEvent statuschanged(m_status)
    End If
    
    Exit Property
    
errhandler:
    m_status = STAT_ERROR
    RaiseEvent statuschanged(m_status)
    ' ---------------------------------------------------------------------------
    ' Raise friendly error to the handler
    ' ---------------------------------------------------------------------------
    If Ambient.UserMode = True Then
        RaiseEvent Error(Err.Number, "CryptoEngine:CryptAlgorithm", Err.Description)
    Else
        Call Err.Raise(Err.Number, "CryptoEngine:CryptAlgorithm", Err.Description)
    End If
End Property

Public Property Get SecurityStatus() As enuSECURITY_STATUS
    SecurityStatus = m_SecStatus
End Property


Public Property Get Status() As Long
    Status = m_status
End Property
Public Property Let Status(ByVal lStatus As Long)
    If lStatus > acuTwofish Or lStatus < acuBlowfish Then
        Err.Raise 1001, "CryptoEngine:Status", "Invalid property value"
    Else
        m_status = lStatus
        RaiseEvent statuschanged(lStatus)
    End If
End Property
Private Sub Reset()
    Set m_Engine = Nothing
    Set m_Blowfish = Nothing
    Set m_DES = Nothing
    Set m_SkipJack = Nothing
    Set m_Gost = Nothing
    Set m_Rijndael = Nothing
    Set m_SkipJack = Nothing
    Set m_TEA = Nothing
    Set m_Twofish = Nothing
End Sub
Private Function SetEngine(ByVal enuCrypt As enuCRYPTO_ALGORITHMS) As Boolean

    SetEngine = False
    
    
    ' Clean Up all
    Call Reset
    
    Select Case enuCrypt
        Case 0: 'enuBlowfish
            Set m_Blowfish = New Blowfish
            Set m_Engine = m_Blowfish
            m_SecStatus = acuHigh
        Case 1: 'enuDES
            Set m_DES = New DES
            Set m_Engine = m_DES
            m_SecStatus = acuLow
        Case 2: 'enuGost
            Set m_Gost = New Gost
            Set m_Engine = m_Gost
            m_SecStatus = acuMedium
        Case 3: ' Rijndael
            Set m_Rijndael = New Rijndael
            Set m_Engine = m_Rijndael
            m_SecStatus = acuHigh
        Case 4: ' Skipjack
            Set m_SkipJack = New Skipjack
            Set m_Engine = m_SkipJack
            m_SecStatus = acuMedium
        Case 5: ' TEA
            Set m_TEA = New TEA
            Set m_Engine = m_TEA
            m_SecStatus = acuMedium
        Case 6: ' Twofish
            Set m_Twofish = New Twofish
            Set m_Engine = m_Twofish
            m_SecStatus = acuHigh
    End Select
    SetEngine = True
 
End Function

Public Function EncryptString(Text As String, Optional Key As String, _
        Optional OutputInHex As Boolean) As String
    
    On Error GoTo errhandler
    If m_status = STAT_BUSY Or m_status = STAT_ERROR Then Exit Function
    EncryptString = m_Engine.EncryptString(Text, Key, OutputInHex)
    m_status = STAT_READY
    RaiseEvent statuschanged(m_status)
    Exit Function
errhandler:
    
    ' ---------------------------------------------------------------------------
    ' Raise friendly error to the handler
    ' ---------------------------------------------------------------------------
    If Ambient.UserMode = True Then
        RaiseEvent Error(Err.Number, "CryptoEngine:EncryptString", Err.Description)
    Else
        Call Err.Raise(Err.Number, "CryptoEngine:EncryptString", Err.Description)
    End If
End Function

Public Function DecryptString(Text As String, Optional Key As String, _
        Optional IsTextInHex As Boolean) As String
    On Error GoTo errhandler
    If m_status = STAT_BUSY Or m_status = STAT_ERROR Then Exit Function
    DecryptString = m_Engine.DecryptString(Text, Key, IsTextInHex)
    m_status = STAT_READY
    RaiseEvent statuschanged(m_status)
    Exit Function
errhandler:
    ' ---------------------------------------------------------------------------
    ' Raise friendly error to the handler
    ' ---------------------------------------------------------------------------
   If Ambient.UserMode = True Then
        RaiseEvent Error(Err.Number, "CryptoEngine:CryptAlgorithm", Err.Description)
    Else
        Call Err.Raise(Err.Number, "CryptoEngine:CryptAlgorithm", Err.Description)
    End If
End Function

Public Function EncryptFile(InFile As String, OutFile As String, _
                Overwrite As Boolean, Optional sKey As String) As Boolean
    On Error GoTo errhandler
    If m_status = STAT_BUSY Or m_status = STAT_ERROR Then Exit Function
    If m_Key <> sKey And Len(sKey) > 0 Then m_Key = sKey
    EncryptFile = m_Engine.EncryptFile(InFile, OutFile, Overwrite, m_Key)
    m_status = STAT_READY
    RaiseEvent statuschanged(m_status)
    Exit Function
errhandler:
    
    ' ---------------------------------------------------------------------------
    ' Raise friendly error to the handler
    ' ---------------------------------------------------------------------------
    If Ambient.UserMode = True Then
        RaiseEvent Error(Err.Number, "CryptoEngine:EncryptFile", Err.Description)
    Else
        Call Err.Raise(Err.Number, "CryptoEngine:EncryptFile", Err.Description)
    End If
End Function

Public Function DecryptFile(InFile As String, OutFile As String, _
        Overwrite As Boolean, Optional Key As String) As Boolean
    On Error GoTo errhandler
    If m_status = STAT_BUSY Or m_status = STAT_ERROR Then Exit Function
    If m_Key <> sKey And Len(sKey) > 0 Then m_Key = sKey
    DecryptFile = m_Engine.DecryptFile(InFile, OutFile, Overwrite, m_Key)
    m_status = STAT_READY
    RaiseEvent statuschanged(m_status)
    Exit Function
errhandler:
    
    ' ---------------------------------------------------------------------------
    ' Raise friendly error to the handler
    ' ---------------------------------------------------------------------------
    If Ambient.UserMode = True Then
        RaiseEvent Error(Err.Number, "CryptoEngine:DecryptFile", Err.Description)
    Else
        Call Err.Raise(Err.Number, "CryptoEngine:DecryptFile", Err.Description)
    End If
End Function

Public Property Let Key(New_Value As String)
On Error GoTo errhandler
    If Len(New_Value) > 0 Then m_Key = New_Value
    Exit Property
errhandler:
    ' ---------------------------------------------------------------------------
    ' Raise friendly error to the handler
    ' ---------------------------------------------------------------------------
    If Ambient.UserMode = True Then
        RaiseEvent Error(Err.Number, "CryptoEngine:Key", Err.Description)
    Else
        Call Err.Raise(Err.Number, "CryptoEngine:Key", Err.Description)
    End If
End Property

Public Property Get Key() As String
    Key = m_Key
End Property
Private Sub SetHash(ByVal eHash As enuHASH_ALGORITHMS)
    If m_status = STAT_BUSY Or m_status = STAT_ERROR Then Exit Sub
    If eHash = acuMD5 Then Set m_Hash = New MD5
    If eHash = acuSHA256 Then Set m_Hash = New SHA256
End Sub

Public Function DigestStringToHex(ByVal strValue As String) As String
    On Error GoTo errhandler
    If m_status = STAT_BUSY Or m_status = STAT_ERROR Then Exit Function
    m_status = STAT_BUSY
    RaiseEvent statuschanged(m_status)
    DigestStringToHex = m_Hash.DigestString(strValue)
    m_status = STAT_READY
    RaiseEvent statuschanged(m_status)
    Exit Function

errhandler:
    m_status = STAT_ERROR
    RaiseEvent statuschanged(m_status)
    ' ---------------------------------------------------------------------------
    ' Raise friendly error to the handler
    ' ---------------------------------------------------------------------------
    If Ambient.UserMode = True Then
        RaiseEvent Error(Err.Number, "CryptoEngine:DigestStringToHex", Err.Description)
    Else
        Call Err.Raise(Err.Number, "CryptoEngine:DigestStringToHex", Err.Description)
    End If
End Function


Public Function DigestFileToHex(ByVal InSource As String) As String
    On Error GoTo errhandler
    If m_status = STAT_BUSY Or m_status = STAT_ERROR Then Exit Function
    m_status = STAT_BUSY
    RaiseEvent statuschanged(m_status)
    DigestFileToHex = m_Hash.DigestFile(InSource)
    RaiseEvent statuschanged(m_status)
    Exit Function

errhandler:
    m_status = STAT_ERROR
    RaiseEvent statuschanged(m_status)
    ' ---------------------------------------------------------------------------
    ' Raise friendly error to the handler
    ' ---------------------------------------------------------------------------
    If Ambient.UserMode = True Then
        RaiseEvent Error(Err.Number, "CryptoEngine:DigestFileToHex", Err.Description)
    Else
        Call Err.Raise(Err.Number, "CryptoEngine:DigestFileToHex", Err.Description)
    End If
End Function

Private Sub Class_Terminate()
    Set m_Engine = Nothing
End Sub


Private Sub m_Blowfish_Process(percent As Long)
    RaiseEvent Process(percent)
End Sub

Private Sub m_Blowfish_StatusChanged(lStatus As Long)
    m_status = lStatus
    RaiseEvent statuschanged(lStatus)
End Sub

Private Sub m_DES_Process(percent As Long)
    RaiseEvent Process(percent)
End Sub

Private Sub m_DES_StatusChanged(lStatus As Long)
    m_status = lStatus
    RaiseEvent statuschanged(lStatus)
End Sub

Private Sub m_Gost_Process(percent As Long)
    RaiseEvent Process(percent)
End Sub

Private Sub m_Gost_StatusChanged(lStatus As Long)
    m_status = lStatus
    RaiseEvent statuschanged(lStatus)
End Sub

Private Sub m_Rijndael_Process(percent As Long)
    RaiseEvent Process(percent)
End Sub

Private Sub m_Rijndael_StatusChanged(lStatus As Long)
    m_status = lStatus
    RaiseEvent statuschanged(lStatus)
End Sub

Private Sub m_SkipJack_Process(percent As Long)
    RaiseEvent Process(percent)
End Sub

Private Sub m_SkipJack_StatusChanged(lStatus As Long)
    m_status = lStatus
    RaiseEvent statuschanged(lStatus)
End Sub

Private Sub m_TEA_Process(percent As Long)
    RaiseEvent Process(percent)
End Sub

Private Sub m_TEA_StatusChanged(lStatus As Long)
    m_status = lStatus
    RaiseEvent statuschanged(lStatus)
End Sub

Private Sub m_Twofish_Process(percent As Long)
    RaiseEvent Process(percent)
End Sub

Private Sub m_Twofish_StatusChanged(lStatus As Long)
    m_status = lStatus
    RaiseEvent statuschanged(lStatus)
End Sub

Private Sub UserControl_Initialize()
    ' Initialize with Blowfish and key = CRYPTOENGINE
    ' Initialize with Hash = SHA256
    HashAlgorithm = acuSHA256
    CryptAlgorithm = acuBlowfish
    Me.Key = "CRYPTOENGINE"

End Sub

Private Sub UserControl_InitProperties()
    HashAlgorithm = acuSHA256
    CryptAlgorithm = acuBlowfish
    Me.Key = "CRYPTOENGINE"
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    HashAlgorithm = PropBag.ReadProperty("HashAlgorithm", acuSHA256)
    CryptAlgorithm = PropBag.ReadProperty("CryptAlgorithm", acuBlowfish)
    m_Key = PropBag.ReadProperty("Key", "CRYPTOENGINE")
End Sub

Private Sub UserControl_Resize()
    Dim R As RECT
    Size DEF_WIDTH, DEF_HEIGHT
    UserControl.ScaleMode = vbPixels
    SetRect R, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    DrawEdge hdc, R, EDGE_RAISED, BF_ADJUST Or BF_RECT
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "HashAlgorithm", m_HashType, acuSHA256
    PropBag.WriteProperty "CryptAlgorithm", m_AlgoType, acuBlowfish
    PropBag.WriteProperty "Key", m_Key, "CRYPTOENGINE"
End Sub
