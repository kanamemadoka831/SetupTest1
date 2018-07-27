VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{78561275-048C-4DB6-937E-38F51FEDAB6E}#1.0#0"; "XP窗体.ocx"
Object = "{F6B9D3BC-3953-4A68-AD1A-BD05206D76A9}#1.0#0"; "hmButton.ocx"
Object = "{307F089F-48D0-466B-9A28-2CBE99958773}#1.0#0"; "MMBT_OSXE.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00F2DED5&
   Caption         =   "Form1"
   ClientHeight    =   3210
   ClientLeft      =   4305
   ClientTop       =   7845
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   11565
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   1800
      Width           =   5295
   End
   Begin 黑马按钮控件.dcButton button3 
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      BackColor       =   -2147483633
      ButtonStyle     =   6
      Caption         =   "请选择包含客户信息的xls文件"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8640
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MMBT_OSXE.MMButton_OSXE Command3 
      Height          =   1455
      Left            =   9360
      TabIndex        =   5
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2566
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "生成注册码文件和注册码"
   End
   Begin Xp窗体.XpCorona XpCorona1 
      Left            =   3000
      Top             =   5160
      _ExtentX        =   4763
      _ExtentY        =   3466
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   1080
      Width           =   5415
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00F2DED5&
      Caption         =   "请输入需要加密的程序（分号相隔）"
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00F2DED5&
      Caption         =   "购买方公司名"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00F2DED5&
      Caption         =   "注册码生成系统"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   8055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
#Const SUPPORT_LEVEL = 0     'Default=0
    Dim excelFileName As String
Dim crc32Table(255) As Long
'Must be equal to SUPPORT_LEVEL in cRijndael
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'An instance of the Class
Private m_Rijndael As New cRijndael
Private Type MD5_CTX
      dwNUMa      As Long
      dwNUMb      As Long
      Buffer(15)  As Byte
      cIN(63)     As Byte
      cDig(15)    As Byte
End Type
Public cn As New ADODB.Connection
Public rs As New ADODB.Recordset
' ==============================
' ≡     API 函 数 声 明      ≡
' ==============================
Private Declare Sub MD5Init Lib "advapi32" (lpContext As MD5_CTX)
Private Declare Sub MD5Final Lib "advapi32" (lpContext As MD5_CTX)
Private Declare Sub MD5Update Lib "advapi32" (lpContext As MD5_CTX, _
                           ByRef lpBuffer As Any, ByVal BufSize As Long)
 
Private stcContext   As MD5_CTX
 Public Function crc32byt(buf() As Byte) As Long
    Dim i As Long, iCRC As Long, lngA As Long, ret As Long
    Dim b() As Byte
    Dim bytT As Byte, bytC As Byte
    b = buf 'StrConv(item, vbFromUnicode)
    iCRC = &HFFFFFFFF
    InitCrc32
    For i = 0 To UBound(b)
        bytC = b(i)
        bytT = (iCRC And &HFF) Xor bytC
        lngA = ((iCRC And &HFFFFFF00) / &H100) And &HFFFFFF
        iCRC = lngA Xor crc32Table(bytT)
    Next
    ret = iCRC Xor &HFFFFFFFF
    crc32byt = ret
End Function



Public Function InitCrc32(Optional ByVal Seed As Long = &HEDB88320, Optional ByVal Precondition As Long = &HFFFFFFFF) As Long
    Dim i As Integer, j As Integer, Crc32 As Long, Temp As Long
    For i = 0 To 255
        Crc32 = i
         For j = 0 To 7
            Temp = ((Crc32 And &HFFFFFFFE) / &H2) And &H7FFFFFFF
            If (Crc32 And &H1) Then Crc32 = Temp Xor Seed Else Crc32 = Temp
        Next
        crc32Table(i) = Crc32
    Next
    InitCrc32 = Precondition
End Function
 
' ==============================
' ≡     通用 函数 & 过程     ≡
' ==============================
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' 功    能：计算一个字符串（ANSI编码）的MD5码
' 入口参数：
'  strText     字符串文本
' 返回参数：   MD5码 （16字节的Byte数组）
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function MD5String(strText As String) As Byte()
      Dim aBuffer()     As Byte
 
   Call MD5Init(stcContext)
   If (Len(strText) > 0) Then
      aBuffer = StrConv(strText, vbFromUnicode)
      Call MD5Update(stcContext, aBuffer(0), UBound(aBuffer) + 1)
   Else
      Call MD5Update(stcContext, 0, 0)
   End If
   Call MD5Final(stcContext)
   MD5String = stcContext.cDig
End Function
 
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' 功    能：计算一个字节流的MD5码
' 入口参数：
'  Buffer      Byte数组
'  size        长度（可选，默认计算整个长度）
' 返回参数：   MD5码 （16字节的Byte数组）
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function md5Bytes(Buffer() As Byte, _
                        Optional ByVal size As Long = -1) As Byte()
      Dim U As Long, pBase   As Long
 
   pBase = LBound(Buffer)
   U = UBound(Buffer) - pBase
   If (-1 = size) Then size = U + 1
   Call MD5Init(stcContext)
   If (-1 = U) Then
      Call MD5Update(stcContext, 0, 0)
   Else
      Call MD5Update(stcContext, Buffer(pBase), size)
   End If
   Call MD5Final(stcContext)
   md5Bytes = stcContext.cDig
End Function
 
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' 功    能：计算一个文件的MD5码
' 入口参数：
'  FileName    磁盘文件名（完整路径）
' 返回参数：   MD5码 （16字节的Byte数组）
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function MD5File(ByVal FileName As String) As Byte()
   Const BUFFERSIZE  As Long = 1024& * 512      ' 缓冲区 512KB
      Dim DataBuff() As Byte
      Dim lFileSize  As Long
      Dim iFn        As Long
 
   On Error GoTo E_Handle_MD5
   If (Len(Dir$(FileName)) = 0) Then Err.Raise 5      '文件不存在
   ReDim DataBuff(BUFFERSIZE - 1)
   iFn = FreeFile()
   Open FileName For Binary As #iFn
   lFileSize = LOF(iFn)
   Call MD5Init(stcContext)
   If (lFileSize = 0) Then
      Call MD5Update(stcContext, 0, 0)
   Else
      Do While (lFileSize > 0)
         Get iFn, , DataBuff
         If (lFileSize > BUFFERSIZE) Then
            Call MD5Update(stcContext, DataBuff(0), BUFFERSIZE)
         Else
            Call MD5Update(stcContext, DataBuff(0), lFileSize)
         End If
         lFileSize = lFileSize - BUFFERSIZE
      Loop
   End If
   Close iFn
   Call MD5Final(stcContext)
E_Handle_MD5:
   MD5File = stcContext.cDig
End Function
 
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' 功    能：获取上次计算的MD5码文本
' 入口参数：   < 无 >
' 返回参数：   MD5码文本字符串（没有MD5数据 返回空串）
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function GetMD5Text() As String
      Dim sResult As String, i&
   If (stcContext.dwNUMa = 0) Then
      sResult = vbNullString
   Else
      sResult = Space$(32)
      For i = 0 To 15
         Mid$(sResult, i + i + 1) = Right$("0" & Hex$(stcContext.cDig(i)), 2)
      Next
   End If
   GetMD5Text = sResult       ' LCase$(sResult) '字母小写
End Function





'Used to display what the program is doing in the Form's caption
'Public Property Let Status(TheStatus As String)
'    If Len(TheStatus) = 0 Then
'        Me.Caption = App.Title
'    Else
'        Me.Caption = App.Title & " - " & TheStatus
'    End If
'    Me.Refresh
'End Property


'Assign TheString to the Text property of TheTextBox if possible.  Otherwise give warning.
'Private Sub DisplayString(TheTextBox As TextBox, ByVal TheString As String)
'    If Len(TheString) < 65536 Then
'        TheTextBox.Text = TheString
'    Else
'        MsgBox "Can not assign a String larger than 64k " & vbCrLf & _
'               "to the Text property of a TextBox control." & vbCrLf & _
'               "If you need to support Strings longer than 64k," & vbCrLf & _
'               "you can use a RichTextBox control instead.", vbInformation
'    End If
'End Sub


'Returns a String containing Hex values of data(0 ... n-1) in groups of k
Private Function HexDisplay(data() As Byte, n As Long, k As Long) As String
    Dim i As Long
    Dim j As Long
    Dim c As Long
    Dim data2() As Byte

    If LBound(data) = 0 Then
        ReDim data2(n * 4 - 1 + ((n - 1) \ k) * 4)
        j = 0
        For i = 0 To n - 1
            If i Mod k = 0 Then
                If i <> 0 Then
                    data2(j) = 32
                    data2(j + 2) = 32
                    j = j + 4
                End If
            End If
            c = data(i) \ 16&
            If c < 10 Then
                data2(j) = c + 48     ' "0"..."9"
            Else
                data2(j) = c + 55     ' "A"..."F"
            End If
            c = data(i) And 15&
            If c < 10 Then
                data2(j + 2) = c + 48 ' "0"..."9"
            Else
                data2(j + 2) = c + 55 ' "A"..."F"
            End If
            j = j + 4
        Next i
Debug.Assert j = UBound(data2) + 1
        HexDisplay = data2
    End If

End Function


'Reverse of HexDisplay.  Given a String containing Hex values, convert to byte array data()
'Returns number of bytes n in data(0 ... n-1)
'Private Function HexDisplayRev(TheString As String, data() As Byte) As Long
'    Dim i As Long
'    Dim j As Long
'    Dim c As Long
'    Dim d As Long
'    Dim n As Long
'    Dim data2() As Byte
'
'    n = 2 * Len(TheString)
'    data2 = TheString
'
'    ReDim data(n \ 4 - 1)
'
'    d = 0
'    i = 0
'    j = 0
'    Do While j < n
'        c = data2(j)
'        Select Case c
'        Case 48 To 57    '"0" ... "9"
'            If d = 0 Then   'high
'                d = c
'            Else            'low
'                data(i) = (c - 48) Or ((d - 48) * 16&)
'                i = i + 1
'                d = 0
'            End If
'        Case 65 To 70   '"A" ... "F"
'            If d = 0 Then   'high
'                d = c - 7
'            Else            'low
'                data(i) = (c - 55) Or ((d - 48) * 16&)
'                i = i + 1
'                d = 0
'            End If
'        Case 97 To 102  '"a" ... "f"
'            If d = 0 Then   'high
'                d = c - 39
'            Else            'low
'                data(i) = (c - 87) Or ((d - 48) * 16&)
'                i = i + 1
'                d = 0
'            End If
'        End Select
'        j = j + 2
'    Loop
'    n = i
'    If n = 0 Then
'        Erase data
'    Else
'        ReDim Preserve data(n - 1)
'    End If
'    HexDisplayRev = n
'
'End Function


'Returns a byte array containing the password in the txtPassword TextBox control.
'If "Plaintext is hex" is checked, and the TextBox contains a Hex value the correct
'length for the current KeySize, the Hex value is used.  Otherwise, ASCII values
'of the txtPassword characters are used.
'Private Function GetPassword() As Byte()
'    Dim data() As Byte
'
'    If Check1.Value = 0 Then
'        data = StrConv(txtPassword.Text, vbFromUnicode)
'        ReDim Preserve data(31)
'    Else
'        If HexDisplayRev(txtPassword.Text, data) <> (cboKeySize.ItemData(cboKeySize.ListIndex) \ 8) Then
'            data = StrConv(txtPassword.Text, vbFromUnicode)
'            ReDim Preserve data(31)
'        End If
'    End If
'    GetPassword = data
'End Function
'
'
'Private Sub cmdEncrypt_Click()
'    Dim pass()        As Byte
'    Dim plaintext()   As Byte
'    Dim ciphertext()  As Byte
'    Dim KeyBits       As Long
'    Dim BlockBits     As Long
'
'    If Len(Text1.Text) = 0 Then
'        MsgBox "No Plaintext"
'    Else
'        If Len(txtPassword.Text) = 0 Then
'            MsgBox "No Password"
'        Else
'            KeyBits = cboKeySize.ItemData(cboKeySize.ListIndex)
'            BlockBits = cboBlockSize.ItemData(cboBlockSize.ListIndex)
'            pass = GetPassword
'
'            Status = "Converting Text"
'            If Check1.Value = 0 Then
'                plaintext = StrConv(Text1.Text, vbFromUnicode)
'            Else
'                If HexDisplayRev(Text1.Text, plaintext) = 0 Then
'                    MsgBox "Text not Hex data"
'                    Status = ""
'                    Exit Sub
'                End If
'            End If
'
'            Status = "Encrypting Data"
'#If SUPPORT_LEVEL Then
'            m_Rijndael.SetCipherKey pass, KeyBits, BlockBits
'            m_Rijndael.ArrayEncrypt plaintext, ciphertext, 0, BlockBits
'#Else
'            m_Rijndael.SetCipherKey pass, KeyBits
'            m_Rijndael.ArrayEncrypt plaintext, ciphertext, 0
'#End If
'            Status = "Converting Text"
'            DisplayString Text1, HexDisplay(ciphertext, UBound(ciphertext) + 1, BlockBits \ 8)
'            Status = ""
'        End If
'    End If
'End Sub
'Private Sub cmdDecrypt_Click()
'    Dim pass()        As Byte
'    Dim plaintext()   As Byte
'    Dim ciphertext()  As Byte
'    Dim KeyBits       As Long
'    Dim BlockBits     As Long
'
'    If Len(Text1.Text) = 0 Then
'        MsgBox "No Ciphertext"
'    Else
'        If Len(txtPassword.Text) = 0 Then
'            MsgBox "No Password"
'        Else
'            KeyBits = cboKeySize.ItemData(cboKeySize.ListIndex)
'            BlockBits = cboBlockSize.ItemData(cboBlockSize.ListIndex)
'            pass = GetPassword
'
'            Status = "Converting Text"
'            If HexDisplayRev(Text1.Text, ciphertext) = 0 Then
'                MsgBox "Text not Hex data"
'                Status = ""
'                Exit Sub
'            End If
'
'            Status = "Decrypting Data"
'#If SUPPORT_LEVEL Then
'            m_Rijndael.SetCipherKey pass, KeyBits, BlockBits
'            If m_Rijndael.ArrayDecrypt(plaintext, ciphertext, 0, BlockBits) <> 0 Then
'                Status = ""
'                Exit Sub
'            End If
'#Else
'            m_Rijndael.SetCipherKey pass, KeyBits
'            If m_Rijndael.ArrayDecrypt(plaintext, ciphertext, 0) <> 0 Then
'                Status = ""
'                Exit Sub
'            End If
'#End If
'            Status = "Converting Text"
'            If Check1.Value = 0 Then
'                DisplayString Text1, StrConv(plaintext, vbUnicode)
'            Else
'                DisplayString Text1, HexDisplay(plaintext, UBound(plaintext) + 1, BlockBits \ 8)
'            End If
'            Status = ""
'        End If
'    End If
'End Sub
'
'
'Private Sub cmdFileEncrypt_Click()
'    Dim FileName  As String
'    Dim FileName2 As String
'    Dim pass()    As Byte
'    Dim KeyBits   As Long
'    Dim BlockBits As Long
'
'    If Len(txtPassword.Text) = 0 Then
'        MsgBox "No Password"
'    Else
'        FileName = FileDialog(Me, False, "File to Encrypt", "*.*|*.*")
'        If Len(FileName) <> 0 Then
'            FileName2 = FileDialog(Me, True, "Save Encrypted Data As ...", "*.dvp|*.dvp|*.*|*.*", FileName & ".dvp")
'            If Len(FileName2) <> 0 Then
'                RidFile FileName2
'                KeyBits = cboKeySize.ItemData(cboKeySize.ListIndex)
'                BlockBits = cboBlockSize.ItemData(cboBlockSize.ListIndex)
'                pass = GetPassword
'
'                Status = "Encrypting File"
'#If SUPPORT_LEVEL Then
'                m_Rijndael.SetCipherKey pass, KeyBits, BlockBits
'                m_Rijndael.FileEncrypt FileName, FileName2, BlockBits
'#Else
'                m_Rijndael.SetCipherKey pass, KeyBits
'                m_Rijndael.FileEncrypt FileName, FileName2
'#End If
'                Status = ""
'            End If
'        End If
'    End If
'End Sub
'Private Sub cmdFileDecrypt_Click()
'    Dim FileName  As String
'    Dim FileName2 As String
'    Dim pass()    As Byte
'    Dim KeyBits   As Long
'    Dim BlockBits As Long
'
'    If Len(txtPassword.Text) = 0 Then
'        MsgBox "No Password"
'    Else
'        FileName = FileDialog(Me, False, "File to Decrypt", "*.dvp|*.dvp|*.*|*.*")
'        If Len(FileName) <> 0 Then
'            If InStrRev(FileName, ".dvp") = Len(FileName) - 3 Then FileName2 = Left$(FileName, Len(FileName) - 4)
'            FileName2 = FileDialog(Me, True, "Save Decrypted Data As ...", "*.*|*.*", FileName2)
'            If Len(FileName2) <> 0 Then
'                RidFile FileName2
'                KeyBits = cboKeySize.ItemData(cboKeySize.ListIndex)
'                BlockBits = cboBlockSize.ItemData(cboBlockSize.ListIndex)
'                pass = GetPassword
'
'                Status = "Decrypting File"
'#If SUPPORT_LEVEL Then
'                m_Rijndael.SetCipherKey pass, KeyBits, BlockBits
'                m_Rijndael.FileDecrypt FileName2, FileName, BlockBits
'#Else
'                m_Rijndael.SetCipherKey pass, KeyBits
'                m_Rijndael.FileDecrypt FileName2, FileName
'#End If
'                Status = ""
'            End If
'        End If
'    End If
'End Sub


'Private Sub chkTerminal_Click()
'    Static Text1FontName As String
'    Static Text1FontBold As Boolean
'    Static Text1FontSize As Long
'
'    If chkTerminal.Value = 0 Then
'        Text1.FontName = Text1FontName
'        Text1.FontBold = Text1FontBold
'        Text1.FontSize = Text1FontSize
'    Else
'        Text1FontName = Text1.FontName
'        Text1FontBold = Text1.FontBold
'        Text1FontSize = Text1.FontSize
'        Text1.FontName = "Terminal"
'    End If
'End Sub

'Private Sub Form_Initialize()
'
'    cboBlockSize.AddItem "128 Bit"
'    cboBlockSize.ItemData(cboBlockSize.NewIndex) = 128
'#If SUPPORT_LEVEL = 0 Then
'    cboBlockSize.Enabled = False
'#Else
'#If SUPPORT_LEVEL = 2 Then
'    cboBlockSize.AddItem "160 Bit"
'    cboBlockSize.ItemData(cboBlockSize.NewIndex) = 160
'    cmdSizeTest.Visible = True
'#End If
'    cboBlockSize.AddItem "192 Bit"
'    cboBlockSize.ItemData(cboBlockSize.NewIndex) = 192
'#If SUPPORT_LEVEL = 2 Then
'    cboBlockSize.AddItem "224 Bit"
'    cboBlockSize.ItemData(cboBlockSize.NewIndex) = 224
'#End If
'    cboBlockSize.AddItem "256 Bit"
'    cboBlockSize.ItemData(cboBlockSize.NewIndex) = 256
'#End If
'    cboKeySize.AddItem "128 Bit"
'    cboKeySize.ItemData(cboKeySize.NewIndex) = 128
'#If SUPPORT_LEVEL = 2 Then
'    cboKeySize.AddItem "160 Bit"
'    cboKeySize.ItemData(cboKeySize.NewIndex) = 160
'#End If
'    cboKeySize.AddItem "192 Bit"
'    cboKeySize.ItemData(cboKeySize.NewIndex) = 192
'#If SUPPORT_LEVEL = 2 Then
'    cboKeySize.AddItem "224 Bit"
'    cboKeySize.ItemData(cboKeySize.NewIndex) = 224
'#End If
'    cboKeySize.AddItem "256 Bit"
'    cboKeySize.ItemData(cboKeySize.NewIndex) = 256
'    cboBlockSize.ListIndex = 0
'    cboKeySize.ListIndex = 0
'    txtPassword = "My Password"
'    Status = ""
'
'End Sub


'COMPLIANCE TESTING
'
'There are many AES and Rijndael Test Vector Files available on the internet so you can
'verify that an implementation is correct.  Below is a simple test that encrypts and
'decrypts one block for each of the 25 combinations of block and key size.  These test
'vectors were created by Dr Brian Gladman.
'
'If the "Plaintext is hex" CheckBox is checked, plaintext is read and written as Hex values,
'just like the ciphertext.  Also, you can enter a Hex value in the txtPassword TextBox.
'To use the "Plaintext is hex" CheckBox, you need to make it visible yourself.  Then you
'can "cut and paste" data directly from known answer test value files.
'
'I've done a reasonable amount of compliance testing, including a few (10,000 iteration) monte
'carlo tests.  I am fairly certain that the class is 100% compliant.  If you find any problems
'or strange behavior, please let me know so it can be corrected.
'
'#If SUPPORT_LEVEL = 2 Then
'Private Sub TestStuff(plaintext As String, passtext As String, ciphertext As String)
'    Dim k As Long
'    Dim p1() As Byte
'    Dim c1() As Byte
'    Dim cdata() As Byte
'    Dim pdata() As Byte
'    Dim pass() As Byte
'    Dim Nk As Long
'    Dim Nb As Long
'    Dim n As Long
'
'    k = HexDisplayRev(passtext, pass)
'    Nk = k \ 4
'    If Nk * 4 <> k Or Nk < 4 Or Nk > 8 Then Exit Sub
'
'    n = HexDisplayRev(plaintext, pdata)
'    Nb = n \ 4
'    If Nb * 4 <> n Or Nb < 4 Or Nb > 8 Then Exit Sub
'
'    If n <> HexDisplayRev(ciphertext, cdata) Then Exit Sub
'
'    m_Rijndael.SetCipherKey pass, Nk * 32, Nb * 32
'    m_Rijndael.ArrayEncrypt pdata, c1, 0, Nb * 32
'    m_Rijndael.ArrayDecrypt p1, cdata, 0, Nb * 32
'
'    Text1.Text = Text1.Text & vbCrLf & "ENCRYPT TEST  " & CStr(Nb * 4) & " byte block, " & CStr(Nk * 4) & " byte key" & vbCrLf
'    Text1.Text = Text1.Text & "KEY:          " & passtext & IIf(UCase$(passtext) = HexDisplay(pass, Nk * 4, Nk * 4), " = ", "<>") & vbCrLf & String(14, 32) & HexDisplay(pass, Nk * 4, Nk * 4) & vbCrLf
'    Text1.Text = Text1.Text & "PLAINTEXT:    " & plaintext & IIf(UCase$(plaintext) = HexDisplay(p1, Nb * 4, Nb * 4), " = ", "<>") & vbCrLf & String(14, 32) & HexDisplay(p1, Nb * 4, Nb * 4) & vbCrLf
'    Text1.Text = Text1.Text & "CIPHERTEXT:   " & ciphertext & IIf(UCase$(ciphertext) = HexDisplay(c1, Nb * 4, Nb * 4), " = ", "<>") & vbCrLf & String(14, 32) & HexDisplay(c1, Nb * 4, Nb * 4) & vbCrLf
'
'End Sub
'Private Sub cmdSizeTest_Click()
'    Text1.Text = ""
'    chkTerminal.Value = 1
'
'    TestStuff "3243f6a8885a308d313198a2e0370734", "2b7e151628aed2a6abf7158809cf4f3c", "3925841d02dc09fbdc118597196a0b32"
'    TestStuff "3243f6a8885a308d313198a2e0370734", "2b7e151628aed2a6abf7158809cf4f3c762e7160", "231d844639b31b412211cfe93712b880"
'    TestStuff "3243f6a8885a308d313198a2e0370734", "2b7e151628aed2a6abf7158809cf4f3c762e7160f38b4da5", "f9fb29aefc384a250340d833b87ebc00"
'    TestStuff "3243f6a8885a308d313198a2e0370734", "2b7e151628aed2a6abf7158809cf4f3c762e7160f38b4da56a784d90", "8faa8fe4dee9eb17caa4797502fc9d3f"
'    TestStuff "3243f6a8885a308d313198a2e0370734", "2b7e151628aed2a6abf7158809cf4f3c762e7160f38b4da56a784d9045190cfe", "1a6e6c2c662e7da6501ffb62bc9e93f3"
'
'    TestStuff "3243f6a8885a308d313198a2e03707344a409382", "2b7e151628aed2a6abf7158809cf4f3c", "16e73aec921314c29df905432bc8968ab64b1f51"
'    TestStuff "3243f6a8885a308d313198a2e03707344a409382", "2b7e151628aed2a6abf7158809cf4f3c762e7160", "0553eb691670dd8a5a5b5addf1aa7450f7a0e587"
'    TestStuff "3243f6a8885a308d313198a2e03707344a409382", "2b7e151628aed2a6abf7158809cf4f3c762e7160f38b4da5", "73cd6f3423036790463aa9e19cfcde894ea16623"
'    TestStuff "3243f6a8885a308d313198a2e03707344a409382", "2b7e151628aed2a6abf7158809cf4f3c762e7160f38b4da56a784d90", "601b5dcd1cf4ece954c740445340bf0afdc048df"
'    TestStuff "3243f6a8885a308d313198a2e03707344a409382", "2b7e151628aed2a6abf7158809cf4f3c762e7160f38b4da56a784d9045190cfe", "579e930b36c1529aa3e86628bacfe146942882cf"
'
'    TestStuff "3243f6a8885a308d313198a2e03707344a4093822299f31d", "2b7e151628aed2a6abf7158809cf4f3c", "b24d275489e82bb8f7375e0d5fcdb1f481757c538b65148a"
'    TestStuff "3243f6a8885a308d313198a2e03707344a4093822299f31d", "2b7e151628aed2a6abf7158809cf4f3c762e7160", "738dae25620d3d3beff4a037a04290d73eb33521a63ea568"
'    TestStuff "3243f6a8885a308d313198a2e03707344a4093822299f31d", "2b7e151628aed2a6abf7158809cf4f3c762e7160f38b4da5", "725ae43b5f3161de806a7c93e0bca93c967ec1ae1b71e1cf"
'    TestStuff "3243f6a8885a308d313198a2e03707344a4093822299f31d", "2b7e151628aed2a6abf7158809cf4f3c762e7160f38b4da56a784d90", "bbfc14180afbf6a36382a061843f0b63e769acdc98769130"
'    TestStuff "3243f6a8885a308d313198a2e03707344a4093822299f31d", "2b7e151628aed2a6abf7158809cf4f3c762e7160f38b4da56a784d9045190cfe", "0ebacf199e3315c2e34b24fcc7c46ef4388aa475d66c194c"
'
'    TestStuff "3243f6a8885a308d313198a2e03707344a4093822299f31d0082efa9", "2b7e151628aed2a6abf7158809cf4f3c", "b0a8f78f6b3c66213f792ffd2a61631f79331407a5e5c8d3793aceb1"
'    TestStuff "3243f6a8885a308d313198a2e03707344a4093822299f31d0082efa9", "2b7e151628aed2a6abf7158809cf4f3c762e7160", "08b99944edfce33a2acb131183ab0168446b2d15e958480010f545e3"
'    TestStuff "3243f6a8885a308d313198a2e03707344a4093822299f31d0082efa9", "2b7e151628aed2a6abf7158809cf4f3c762e7160f38b4da5", "be4c597d8f7efe22a2f7e5b1938e2564d452a5bfe72399c7af1101e2"
'    TestStuff "3243f6a8885a308d313198a2e03707344a4093822299f31d0082efa9", "2b7e151628aed2a6abf7158809cf4f3c762e7160f38b4da56a784d90", "ef529598ecbce297811b49bbed2c33bbe1241d6e1a833dbe119569e8"
'    TestStuff "3243f6a8885a308d313198a2e03707344a4093822299f31d0082efa9", "2b7e151628aed2a6abf7158809cf4f3c762e7160f38b4da56a784d9045190cfe", "02fafc200176ed05deb8edb82a3555b0b10d47a388dfd59cab2f6c11"
'
'    TestStuff "3243f6a8885a308d313198a2e03707344a4093822299f31d0082efa98ec4e6c8", "2b7e151628aed2a6abf7158809cf4f3c", "7d15479076b69a46ffb3b3beae97ad8313f622f67fedb487de9f06b9ed9c8f19"
'    TestStuff "3243f6a8885a308d313198a2e03707344a4093822299f31d0082efa98ec4e6c8", "2b7e151628aed2a6abf7158809cf4f3c762e7160", "514f93fb296b5ad16aa7df8b577abcbd484decacccc7fb1f18dc567309ceeffd"
'    TestStuff "3243f6a8885a308d313198a2e03707344a4093822299f31d0082efa98ec4e6c8", "2b7e151628aed2a6abf7158809cf4f3c762e7160f38b4da5", "5d7101727bb25781bf6715b0e6955282b9610e23a43c2eb062699f0ebf5887b2"
'    TestStuff "3243f6a8885a308d313198a2e03707344a4093822299f31d0082efa98ec4e6c8", "2b7e151628aed2a6abf7158809cf4f3c762e7160f38b4da56a784d90", "d56c5a63627432579e1dd308b2c8f157b40a4bfb56fea1377b25d3ed3d6dbf80"
'    TestStuff "3243f6a8885a308d313198a2e03707344a4093822299f31d0082efa98ec4e6c8", "2b7e151628aed2a6abf7158809cf4f3c762e7160f38b4da56a784d9045190cfe", "a49406115dfb30a40418aafa4869b7c6a886ff31602a7dd19c889dc64f7e4e7a"
'
'End Sub
'#End If



Public Sub connOpen()
' 定义数据库连接字符串变量
Dim strCn As String
 ' 定义数据库连接参数变量
Dim db_host As String
 Dim db_user As String
 Dim db_pass As String
 Dim db_data As String
 ' 初始化数据库连接变量
db_host = "localhost"
 db_user = "root"
db_pass = "Kanamemadoka_831"
db_data = "test"
'创建连接字符串
strCn = "DRIVER={MySQL ODBC 5.1 DRIVER};" & _
 "SERVER=" & db_host & ";" & _
 "DATABASE=" & db_data & ";" & _
 "UID=" & db_user & ";PWD=" & db_pass & ";" & _
 "OPTION=3;stmt=SET NAMES GB2312"
'MySql ODBC的连接参数如下:
'userMySql用户名默认值:ODBC(on Windows)
'serverMySql服务器地址默认值:localhost
'database数据库名
'option连接的工作方式默认值:0
'port连接端口默认值:3306
'stmt一段声明
'password用户密码

'注意:这里的stmt=SET NAMES GB2312
'指定数据库编码方式,中文操作系统需设置成GB2312

'连接数据库
cn.Open strCn
 ' 设置该属性, 使 recordcount 和 absolutepage 属性可用
cn.CursorLocation = adUseClient
End Sub

Private Sub Command1_Click()

'Dim timeEncrypt As String
'MsgBox (time)
'MsgBox (strEncrypt(time, machineCode))
'timeEncrypt = strEncrypt(time, machineCode)
'MsgBox (strDecrypt(timeEncrypt, machineCode))

'Dim md5Code As String
'Dim md5Bytes() As Byte
'md5Bytes = MD5String(time & machineCode)
'md5Code = GetMD5Text()
'Dim finalCodeCal As String
'finalCodeCal = timeEncrypt & md5Code
End Sub

Private Sub button3_Click()
    CommonDialog1.ShowOpen
    
     excelFileName = CommonDialog1.FileName
     Text2.text = excelFileName
End Sub

Private Sub Command3_Click()
Dim time As String
Dim timeplus As String
time = Format(Now(), "yyyy/MM/dd HH:mm:ss")
Dim s As Integer
s = 60
timeplus = DateAdd("s", s, time)
timeplus = Format(timeplus, "yyyy/MM/dd HH:mm:ss")
time = DelComma(time)
time = DelSlash(time)
time = DelSpace(time)

timeplus = DelComma(timeplus)
timeplus = DelSlash(timeplus)
timeplus = DelSpace(timeplus)

Dim ranNum As Integer
ranNum = Rnd * (10000 - 0 + 1) + 0
Dim year As String
Dim month As String
Dim DH As String
Dim MS As String
Dim timeOrg As String
timeOrg = time
year = Left(time, 4)
month = Right(Left(time, 6), 2)
DH = Right(Left(time, 10), 4)
MS = Right(time, 4)
Dim ranStr As String
ranStr = CStr(ranNum)
time = year + Left(ranStr, 3) + DH + Right(ranStr, 2) + month + MS

Dim yearplus As String
Dim monthplus As String
Dim DHplus As String
Dim MSplus As String
Dim timeOrgplus As String
timeOrgplus = timeplus
yearplus = Left(timeplus, 4)
monthplus = Right(Left(timeplus, 6), 2)
DHplus = Right(Left(timeplus, 10), 4)
MSplus = Right(timeplus, 4)
Dim ranStrplus As String
ranStrplus = CStr(ranNum)
timeplus = yearplus + Left(ranStrplus, 3) + DHplus + Right(ranStrplus, 2) + monthplus + MSplus

Dim xl As Object
Set xl = New Excel.Application
xl.Visible = True
xl.Workbooks.Add
xl.Workbooks.Open (excelFileName)
Dim rng As Range
Dim jump As Integer, j As Integer, lengthOfRows As Integer
Set rng = xl.ActiveSheet.UsedRange
lengthOfRows = rng.Rows.Count
Dim lengthOfCodes As Integer
lengthOfCodes = lengthOfRows - 2
Dim machineFromExcel() As String
ReDim machineFromExcel(lengthOfCodes)
For jump = 2 To lengthOfRows
    machineFromExcel(jump - 2) = xl.ActiveSheet.Cells(jump, 2)
'    MsgBox machineFromExcel(jump - 2)
Next
Dim machineFromExcels As String
machineFromExcels = Join(machineFromExcel, ";")


Dim machineCodeAll As String
Dim machineCode() As String
Dim companyName As String
companyName = Text1.text
'machineCodeAll = Text2.text
machineCode = Split(machineFromExcels, ";")
Dim Length As Integer
Length = UBound(machineCode) - LBound(machineCode) + 1
Dim finalCodePlus() As String
ReDim finalCodePlus(Length)
Dim finalCode() As String
ReDim finalCode(Length)
Dim sqls() As String
ReDim sqls(Length)
Dim i As Integer
If Length <> 1 Then
For i = LBound(machineCode) To UBound(machineCode)
    MD5String (time & machineCode(i))
    finalCode(i) = strEncrypt(time, machineCode(i)) & GetMD5Text() & vbCrLf
    finalCode(i) = DelSpace(finalCode(i))
    MD5String (timeplus & machineCode(i))
    finalCodePlus(i) = strEncrypt(timeplus, machineCode(i)) & GetMD5Text()
    finalCodePlus(i) = DelSpace(finalCodePlus(i))
    'isValuedCode (finalCodePlus(i))
    sqls(i) = "INSERT INTO Buyinformation(company_id,time_table,machine_id,Filecode,code) VALUES(""" & companyName & """,""" & CStr(timeOrg) & """,""" & machineCode(i) & """,""" & finalCodePlus(i) & """,""" & finalCode(i) & """) "
    cn.Execute sqls(i)
    xl.ActiveSheet.Cells(i + 2, 3) = finalCode(i)
    Open App.Path & "\enregistrait" & xl.ActiveSheet.Cells(i + 2, 1) & ".txt" For Output As #1
    Print #1, finalCodePlus(i)
    Close #1
    xl.ActiveSheet.Cells(i + 2, 4) = "\enregistrait" & xl.ActiveSheet.Cells(i + 2, 1) & ".txt"
Next
End If
'MsgBox "youyouyou"
xl.Visible = False
xl.ActiveWorkbook.Save
xl.Quit
'If length = 1 Then
'    MD5String (time & machineCode(i))
'    finalCode(i) = strEncrypt(time, machineCode(i)) & GetMD5Text() & vbCrLf
'    finalCode(i) = DelSpace(finalCode(i))
'    MD5String (timeplus & machineCode(i))
'    finalCodePlus(i) = strEncrypt(timeplus, machineCode(i)) & GetMD5Text()
'    finalCodePlus(i) = DelSpace(finalCodePlus(i))
'    isValuedCode (finalCodePlus(i))
'    sqls(i) = "INSERT INTO Buyinformation(company_id,time_table,machine_id,Filecode,code) VALUES(""" & companyName & """,""" & CStr(timeOrg) & """,""" & machineCode(i) & """,""" & finalCodePlus(i) & """,""" & finalCode(i) & """) "
'    cn.Execute sqls(i)
'    Open App.Path & "\enregistrait.txt" For Output As #1
'    Print #1, finalCodePlus(i)
'    Close #1
'End If

'Dim finalCodes As String
'For i = LBound(finalCode) To UBound(finalCode)
'    finalCodes = finalCodes & finalCode(i) & vbCrLf
'Next
'Text3.text = finalCodes

Dim exeNamesAll As String
exeNamesAll = Text4.text
Dim exeName() As String
exeName = Split(exeNamesAll, ";")
Dim fBytes() As Byte
Dim length1 As Long
Dim FileStr As String
Dim NewCRC As String
Dim NewStr As String
Call deleteOrigFiles(App.Path & "\Untitled.suf")
Call AddNewExes(App.Path & "\Untitled.suf", exeName)
Call deleteOrigIFS(App.Path & "\Untitled.suf")
Call addexeIFS(App.Path & "\Untitled.suf", exeName)
Shell "C:\Program Files (x86)\Setup Factory 9\SUFDesign.exe " & App.Path & """\Untitled.suf"""
Sleep "2000"
SendKeys "{F7}", True
SendKeys "{Enter}", True
SendKeys "{Enter}", True
SendKeys "{Enter}", True
Sleep "30000"
SendKeys "{Enter}", True
SendKeys "%{F4}", True
End Sub

Private Sub Form_Load()
Dim sql As String
Dim con As String
'连接数据库
connOpen
'创建SQL语句
'假设在table表中存在字段name,其中包含一个值"MyName"
sql = "CREATE TABLE IF NOT EXISTS BuyInformation(company_id VARCHAR(100) not NULL,time_table VARCHAR(100) not NULL,machine_id VARCHAR(100) not NULL,Filecode VARCHAR(100) not NULL,code VARCHAR(100) not NULL)"
con = "DRIVER={MySQL ODBC 5.1 DRIVER};" & "SERVER=" & "localhost" & ";" & "DATABASE=" & "test" & ";" & "UID=" & "root" & ";PWD=" & "Kanamemadoka_831" & ";" & "OPTION=3;stmt=SET NAMES GB2312"
'rs.Open sql, con
cn.Execute sql

'MsgBox将显示"MyName"
'垃圾回收
'rs.Close
'cn.Close
'不再需要数据集和连接时,比如退出程序时
'Set rs = Nothing
'Set cn = Nothing
End Sub
'Private Function isValuedCode(text As String) As Boolean
'Dim machine As String
'Dim mo
'machine = ""
'Dim mc
'Set mc = GetObject("Winmgmts:").InstancesOf("Win32_NetworkAdapterConfiguration")
'For Each mo In mc
'If mo.IPEnabled = True Then
'machine = machine & DelComma(mo.MacAddress)
'Exit For
'End If
'Next
'
'Dim HDid, moc
'Set moc = GetObject("Winmgmts:").InstancesOf("Win32_DiskDrive")
'For Each mo In moc
'HDid = mo.SerialNumber
'machine = machine & LTrim(HDid)
'Next
'
'Dim cpuInfo
'cpuInfo = ""
'Set moc = GetObject("Winmgmts:").InstancesOf("Win32_Processor")
'For Each mo In moc
'cpuInfo = CStr(mo.ProcessorId)
'machine = machine & cpuInfo
'Next
'
'Dim biosId
'biosId = ""
'Set moc = GetObject("Winmgmts:").InstancesOf("Win32_BIOS")
'Dim serNum
'For Each mo In moc
'serNum = CStr(mo.SerialNumber)
'machine = machine & serNum
'Next
'Dim AESString As String
'Dim timeCode As String
'Dim MDString As String
'AESString = Left(text, 64)
'MDString = Right(text, Len(text) - 64)
'timeCode = strDecrypt(AESString, Text2.text)
'Dim timeCode1 As Integer
'Dim timeCode2 As Integer
'Dim timeCode3 As Integer
'Dim timeCode4 As Integer
'Dim md5Bytes() As Byte
'Dim md5Code As String
'md5Bytes = MD5String(time & machine)
'md5Code = GetMD5Text()
'If md5Code <> MDString Then
'    isValuedCode = False
'Else
'    timeCode1 = CInt(Left(timeCode, 4))
'    timeCode3 = CInt(Right(Left(timeCode, 11), 4))
'    timeCode2 = CInt(Left(Right(timeCode, 6), 2))
'    timeCode4 = CInt(Right(timeCode, 4))
'    If (timeCode1 > 2100) Then
'        isValuedCode = False
'    Else
'        If (timeCode2 > 12) Then
'            isValuedCode = False
'        Else
'            If (timeCode2 < 1) Then
'                isValuedCode = False
'            Else
'                If (timeCode3 > 3124) Then
'                    isValuedCode = False
'                Else
'                    If (timeCode3 < 0) Then
'                        isValuedCode = False
'                    Else
'                        If (timeCode4 > 6060) Then
'                            isValuedCode = False
'                        Else
'                            If (timeCode4 < 0) Then
'                                isValuedCode = False
'                            End If
'                        End If
'                    End If
'                End If
'            End If
'        End If
'    End If
'End If
'isValuedCode = True
'
'End Function
Private Function strEncrypt(text As String, password As String) As String
    Dim pass()        As Byte
    Dim plaintext()   As Byte
    Dim ciphertext()  As Byte
    Dim KeyBits       As Long
    Dim BlockBits     As Long

    If Len(text) = 0 Then
        MsgBox "No Plaintext"
    Else
        If Len(password) = 0 Then
            MsgBox "No Password"
        Else
            'KeyBits = cboKeySize.ItemData(cboKeySize.ListIndex)
            KeyBits = 128
            'BlockBits = cboBlockSize.ItemData(cboBlockSize.ListIndex)
            BlockBits = 128
            pass = password

            'Status = "Converting Text"
'            If Check1.Value = 0 Then
                plaintext = StrConv(text, vbFromUnicode)
'            Else
'                If HexDisplayRev(text, plaintext) = 0 Then
'                    MsgBox "Text not Hex data"
'                    Status = ""
'                    Exit Function
'                End If
'            End If

           ' Status = "Encrypting Data"
#If SUPPORT_LEVEL Then
            m_Rijndael.SetCipherKey pass, KeyBits, BlockBits
            m_Rijndael.ArrayEncrypt plaintext, ciphertext, 0, BlockBits
#Else
            m_Rijndael.SetCipherKey pass, KeyBits
            m_Rijndael.ArrayEncrypt plaintext, ciphertext, 0
#End If
           ' Status = "Converting Text"
'            DisplayString Text1, HexDisplay(ciphertext, UBound(ciphertext) + 1, BlockBits \ 8)
            strEncrypt = HexDisplay(ciphertext, UBound(ciphertext) + 1, BlockBits \ 8)
            'Status = ""
        End If
    End If
End Function
Function DelComma(str As String)
    DelComma = Join(Split(str, ":"), "")
End Function
Function DelSpace(str As String)
    DelSpace = Join(Split(str, " "), "")
End Function
Function DelSlash(str As String)
    DelSlash = Join(Split(str, "/"), "")
End Function
Private Function ChangeLine(strFile As String, RLine As Long, NewStr As String)
Dim s As String, n As String, i As Long
i = 1

'//打开源文件
Open strFile For Input As #1
Do Until EOF(1)
    Line Input #1, s
    If RLine = i Then '如果是指定的行数就进行下面的操作
            s = NewStr
            n = n & s & vbCrLf '将空字符串赋给变量n,以保持源文件的行数
        Else    '如果不是指定的行数,就将s的内容赋给变量n 以存储数据
        n = n & s & vbCrLf   '将s的内容赋给n 并以一个回车符号结束....
    End If
    i = i + 1
Loop
Close #1

   '//写入新文件,如果和源文件同名则会覆盖源文件
Open strFile For Output As #2
Print #2, n '将n变量里的数据写入新文件
Close #2
End Function
Private Function strDecrypt(text As String, password As String) As String
    Dim pass()        As Byte
    Dim plaintext()   As Byte
    Dim ciphertext()  As Byte
    Dim KeyBits       As Long
    Dim BlockBits     As Long

    If Len(text) = 0 Then
    Else
        If Len(password) = 0 Then
        Else
            'KeyBits = cboKeySize.ItemData(cboKeySize.ListIndex)
            KeyBits = 128
            'BlockBits = cboBlockSize.ItemData(cboBlockSize.ListIndex)
            BlockBits = 128
            pass = password

'            Status = "Converting Text"
            If HexDisplayRev(text, ciphertext) = 0 Then
'                Status = ""
                Exit Function
            End If

'            Status = "Decrypting Data"
#If SUPPORT_LEVEL Then
            m_Rijndael.SetCipherKey pass, KeyBits, BlockBits
            If m_Rijndael.ArrayDecrypt(plaintext, ciphertext, 0, BlockBits) <> 0 Then
'                Status = ""
                Exit Function
            End If
#Else
            m_Rijndael.SetCipherKey pass, KeyBits
            If m_Rijndael.ArrayDecrypt(plaintext, ciphertext, 0) <> 0 Then
'                Status = ""
                Exit Function
            End If
#End If
'            Status = "Converting Text"
'            If Check1.Value = 0 Then
                'DisplayString Text1, StrConv(plaintext, vbUnicode)
                strDecrypt = StrConv(plaintext, vbUnicode)
'            Else
'                'DisplayString Text1, HexDisplay(plaintext, UBound(plaintext) + 1, BlockBits \ 8)
'                strDecrypt = HexDisplay(plaintext, UBound(plaintext) + 1, BlockBits \ 8)
'            End If
'            Status = ""
        End If
    End If
End Function
Private Function HexDisplayRev(TheString As String, data() As Byte) As Long
    Dim i As Long
    Dim j As Long
    Dim c As Long
    Dim d As Long
    Dim n As Long
    Dim data2() As Byte

    n = 2 * Len(TheString)
    data2 = TheString

    ReDim data(n \ 4 - 1)

    d = 0
    i = 0
    j = 0
    Do While j < n
        c = data2(j)
        Select Case c
        Case 48 To 57    '"0" ... "9"
            If d = 0 Then   'high
                d = c
            Else            'low
                data(i) = (c - 48) Or ((d - 48) * 16&)
                i = i + 1
                d = 0
            End If
        Case 65 To 70   '"A" ... "F"
            If d = 0 Then   'high
                d = c - 7
            Else            'low
                data(i) = (c - 55) Or ((d - 48) * 16&)
                i = i + 1
                d = 0
            End If
        Case 97 To 102  '"a" ... "f"
            If d = 0 Then   'high
                d = c - 39
            Else            'low
                data(i) = (c - 87) Or ((d - 48) * 16&)
                i = i + 1
                d = 0
            End If
        End Select
        j = j + 2
    Loop
    n = i
    If n = 0 Then
        Erase data
    Else
        ReDim Preserve data(n - 1)
    End If
    HexDisplayRev = n

End Function
Private Function AddNewExes(strFile As String, exeNames() As String)
Open strFile For Binary As #1    '文本文件名自己改
Dim s() As String
s = Split(Input(LOF(1), #1), vbCrLf)
s(5) = s(5) & vbCrLf  's(0)是第一行，s(1)是第二行，依此类推
Dim n As String, i As Integer
Dim number As Integer
number = UBound(exeNames) - LBound(exeNames) + 1
Dim Addons As String
Addons = ""
Dim Adds(0 To 67) As String
For i = LBound(exeNames) To UBound(exeNames)
    Adds(0) = "<FileData>" & vbCrLf
    Adds(1) = "<FldRef>0</FldRef>" & vbCrLf
    Adds(2) = "<FullName>" & App.Path & "\" & exeNames(i) & ".exe</FullName>" & vbCrLf
    Adds(3) = "<FileName>" & exeNames(i) & ".exe" & "</FileName>" & vbCrLf
    Adds(4) = "<Source>" & App.Path & "</Source>" & vbCrLf
    Adds(5) = "<Ext>exe</Ext>" & vbCrLf
    Adds(6) = "<RTSource>Archive</RTSource>" & vbCrLf
    Adds(7) = "<Desc/>" & vbCrLf
    Adds(8) = "<Recurse>1</Recurse>" & vbCrLf
    Adds(9) = "<MatchMode>0</MatchMode>" & vbCrLf
    Adds(10) = "<Dest>%AppFolder%</Dest>" & vbCrLf
    Adds(11) = "<Overwrite>1</Overwrite>" & vbCrLf
    Adds(12) = "<Backup>0</Backup>" & vbCrLf
    Adds(13) = "<Protect>0</Protect>" & vbCrLf
    Adds(14) = "<InstallOrder>1000</InstallOrder>" & vbCrLf
    Adds(15) = "<SCStartRoot>0</SCStartRoot>" & vbCrLf
    Adds(16) = "<SCStartProgs>0</SCStartProgs>" & vbCrLf
    Adds(17) = "<SCAppFld>1</SCAppFld>" & vbCrLf
    Adds(18) = "<SCStartup>0</SCStartup>" & vbCrLf
    Adds(19) = "<SCDesk>0</SCDesk>" & vbCrLf
    Adds(20) = "<SCQLaunch>0</SCQLaunch>" & vbCrLf
    Adds(21) = "<SCCust>0</SCCust>" & vbCrLf
    Adds(22) = "<CustSCPath/>" & vbCrLf
    Adds(23) = "<SCDesc>" & exeNames(i) & "</SCDesc>" & vbCrLf
    Adds(24) = "<SCComment/>" & vbCrLf
    Adds(25) = "<SCArgs/>" & vbCrLf
    Adds(26) = "<SCWork/>" & vbCrLf
    Adds(27) = "<UseExtIco>0</UseExtIco>" & vbCrLf
    Adds(28) = "<IcoFN/>" & vbCrLf
    Adds(29) = "<IcoIdx>0</IcoIdx>" & vbCrLf
    Adds(30) = "<IcoShowMd>0</IcoShowMd>" & vbCrLf
    Adds(31) = "<IcoHK>0</IcoHK>" & vbCrLf
    Adds(32) = "<RegTTF>0</RegTTF>" & vbCrLf
    Adds(33) = "<TTFName/>" & vbCrLf
    Adds(34) = "<RegOCX>0</RegOCX>" & vbCrLf
    Adds(35) = "<RegTLB>0</RegTLB>" & vbCrLf
    Adds(36) = "<SupInUse>0</SupInUse>" & vbCrLf
    Adds(37) = "<Compress>1</Compress>" & vbCrLf
    Adds(38) = "<UseOrigAttr>1</UseOrigAttr>" & vbCrLf
    Adds(39) = "<Attr>0</Attr>" & vbCrLf
    Adds(40) = "<NoCRC>0</NoCRC>" & vbCrLf
    Adds(41) = "<NoRemove>0</NoRemove>" & vbCrLf
    Adds(42) = "<Shared>0</Shared>" & vbCrLf
    Adds(43) = "<OSCond>" & vbCrLf
    Adds(44) = "<OS>32768</OS>" & vbCrLf
    Adds(45) = "<OS>65535</OS>" & vbCrLf
    Adds(46) = "<OS>65535</OS>" & vbCrLf
    Adds(47) = "<OS>65535</OS>" & vbCrLf
    Adds(48) = "<OS>65535</OS>" & vbCrLf
    Adds(49) = "<OS>65535</OS>" & vbCrLf
    Adds(50) = "<OS>65535</OS>" & vbCrLf
    Adds(51) = "<OS>65535</OS>" & vbCrLf
    Adds(52) = "<OS>65535</OS>" & vbCrLf
    Adds(53) = "<OS>65535</OS>" & vbCrLf
    Adds(54) = "<OS>65535</OS>" & vbCrLf
    Adds(55) = "<OS>65535</OS>" & vbCrLf
    Adds(56) = "</OSCond>" & vbCrLf
    Adds(57) = "<RTCond/>" & vbCrLf
    Adds(58) = "<BuildConfigs>" & vbCrLf
    Adds(59) = "<Cfg>All</Cfg>" & vbCrLf
    Adds(60) = "</BuildConfigs>" & vbCrLf
    Adds(61) = "<Package>None</Package>" & vbCrLf
    Adds(62) = "<Packages/>" & vbCrLf
    Adds(63) = "<Notes/>" & vbCrLf
    Adds(64) = "<CompSize>0</CompSize>" & vbCrLf
    Adds(65) = "<CRC>0</CRC>" & vbCrLf
    Adds(66) = "<StoreOnly>0</StoreOnly>" & vbCrLf
    Adds(67) = "</FileData>" & vbCrLf
    Addons = Addons & Join(Adds)
Next
s(5) = s(5) & Addons
Put #1, 1, Join(s, vbCrLf)
Close #1
End Function
Private Function deleteOrigFiles(StrFileName As String)
    Open StrFileName For Binary As #1    '文本文件名自己改
    Dim final As String
    Dim s() As String
    Dim ColNum As Integer
    Dim cols As String
    s = Split(Input(LOF(1), #1), vbCrLf)
    Close #1
    Dim Length As Integer
    Length = UBound(s) - LBound(s) + 1
    Dim i As Integer
    For i = 0 To Length - 1
        If InStr(1, s(i), "/ArchiveFiles") > 0 Then
        ColNum = i
        End If
    Next
    'MsgBox ColNum
    cols = ""
    For i = ColNum To Length - 1
        cols = cols & s(i) & vbCrLf
    Next
    final = ""
    final = final & s(0) & vbCrLf
    final = final & s(1) & vbCrLf
    final = final & s(2) & vbCrLf
    final = final & s(3) & vbCrLf
    final = final & s(4) & vbCrLf
    final = final & s(5) & vbCrLf
    final = final & cols
    Open App.Path & "\untitled.suf" For Output As #1
    Print #1, final
    Close #1
End Function
Private Function deleteOrigIFS(StrFileName As String)
    Open StrFileName For Binary As #1
    Dim final As String
    Dim s() As String
    Dim query1 As String
    Dim query2 As String
    Dim Length As Integer
    s = Split(Input(LOF(1), #1), vbCrLf)
    Length = UBound(s) - LBound(s) + 1
    query1 = "strSerial == serialNumber"
    query2 = "Screen.Jump(""Ready"
    Close #1
    Dim queryCo1 As Integer
    Dim queryCo2 As Integer
    Dim i As Integer
    For i = 0 To Length - 1
        If InStr(1, s(i), query1) > 0 Then
        queryCo1 = i
        End If
    Next
    'MsgBox queryCo1
    For i = 0 To Length - 1
        If InStr(s(i), query2) > 0 Then
        queryCo2 = i
        End If
    Next
    'MsgBox queryCo2
    'MsgBox Length
    final = ""
    For i = 0 To queryCo1
        final = final & s(i) & vbCrLf
    Next
    For i = queryCo2 To Length - 1
        final = final & s(i) & vbCrLf
    Next
    Open StrFileName For Output As #1
    Print #1, final
    Close #1
End Function
Private Function addexeIFS(StrFileName As String, exeNames() As String)
    Open StrFileName For Binary As #1
    Dim final As String
    Dim query1 As String
    Dim query2 As String
    Dim Length As Long
    Dim length1 As Long
    Dim NewStr As String
    Dim fBytes() As Byte
    Dim s() As String
    s = Split(Input(LOF(1), #1), vbCrLf)
    Close #1
    Length = UBound(s) - LBound(s) + 1
    query1 = "strSerial == serialNumber"
    query2 = "Screen.Jump(""Ready"
    Dim queryCo1 As Integer
    Dim queryCo2 As Integer
    Dim i As Integer
    For i = 0 To Length - 1
        If InStr(1, s(i), query1) > 0 Then
        queryCo1 = i
        End If
    Next
    For i = 0 To Length - 1
        If InStr(1, s(i), query2) > 0 Then
        queryCo2 = i
        End If
    Next
    final = ""
    For i = 0 To queryCo1
        final = final & s(i) & vbCrLf
    Next
    NewStr = ""
    Dim NewCRC As String
    NewStr = NewStr & "Registry.SetValue(HKEY_LOCAL_MACHINE,""SOFTWARE\\JiCheng"",""isRegister"",strSerial,REG_SZ );" & vbCrLf
    NewStr = NewStr & "Registry.SetValue(HKEY_LOCAL_MACHINE,""SOFTWARE\\gnehCiJ"",""isRegister"",strSerial,REG_SZ );" & vbCrLf
    NewStr = NewStr & "Registry.SetValue(HKEY_CURRENT_USER,""SOFTWARE\\JiCheng"",""isRegister"",strSerial,REG_SZ );" & vbCrLf
    NewStr = NewStr & "Registry.SetValue(HKEY_CURRENT_USER,""SOFTWARE\\gnehcij"",""isRegister"",strSerial,REG_SZ );" & vbCrLf
    For i = 0 To UBound(exeNames) - LBound(exeNames)
    'MsgBox (exeNames(i) & ".exe")
        Open App.Path & "\" & exeNames(i) & ".exe" For Binary As #2
        length1 = LOF(2)
        ReDim fBytes(LOF(2))
        Get #2, , fBytes
        Close #2
        NewCRC = Right("00" & Hex(crc32byt(fBytes)), 8)
        NewStr = NewStr & "Registry.SetValue(HKEY_LOCAL_MACHINE,""SOFTWARE\\JiCheng\\" & exeNames(i) & """,""MD5"",""" & NewCRC & """,REG_SZ );" & vbCrLf
        NewStr = NewStr & "Registry.SetValue(HKEY_LOCAL_MACHINE,""SOFTWARE\\gnehCiJ\\" & exeNames(i) & """,""MD5"",""" & NewCRC & """,REG_SZ );" & vbCrLf
        NewStr = NewStr & "Registry.SetValue(HKEY_CURRENT_USER,""SOFTWARE\\JiCheng\\" & exeNames(i) & """,""MD5"",""" & NewCRC & """,REG_SZ );" & vbCrLf
        NewStr = NewStr & "Registry.SetValue(HKEY_CURRENT_USER,""SOFTWARE\\gnehcij\\" & exeNames(i) & """,""MD5"",""" & NewCRC & """,REG_SZ );" & vbCrLf
    Next
    final = final & NewStr
    For i = queryCo2 To Length - 1
        final = final & s(i) & vbCrLf
    Next
    Open StrFileName For Output As #1
    Print #1, final
    Close #1
End Function

