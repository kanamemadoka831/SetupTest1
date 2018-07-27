Attribute VB_Name = "check"
Option Explicit
#Const SUPPORT_LEVEL = 0     'Default=0
'Must be equal to SUPPORT_LEVEL in cRijndael
Dim crc32Table(255) As Long
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
Private Declare Sub MD5Init Lib "advapi32" (lpContext As MD5_CTX)
Private Declare Sub MD5Final Lib "advapi32" (lpContext As MD5_CTX)
Private Declare Sub MD5Update Lib "advapi32" (lpContext As MD5_CTX, _
                           ByRef lpBuffer As Any, ByVal BufSize As Long)

 
Private stcContext   As MD5_CTX
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
Private Function getTime(str As String) As Integer
Dim machine As String
Dim mo
machine = ""
Dim mc
Set mc = GetObject("Winmgmts:").InstancesOf("Win32_NetworkAdapterConfiguration")
For Each mo In mc
If mo.IPEnabled = True Then
machine = machine & DelComma(mo.MacAddress)
Exit For
End If
Next

Dim HDid, moc
Set moc = GetObject("Winmgmts:").InstancesOf("Win32_DiskDrive")
For Each mo In moc
HDid = mo.SerialNumber
machine = machine & LTrim(HDid)
Next

Dim cpuInfo
cpuInfo = ""
Set moc = GetObject("Winmgmts:").InstancesOf("Win32_Processor")
For Each mo In moc
cpuInfo = CStr(mo.ProcessorId)
machine = machine & cpuInfo
Next

Dim biosId
biosId = ""
Set moc = GetObject("Winmgmts:").InstancesOf("Win32_BIOS")
Dim serNum
For Each mo In moc
serNum = CStr(mo.SerialNumber)
machine = machine & serNum
Next
str = Left(str, 64)
Dim timeCode As String
timeCode = strDecrypt(str, machine)
Dim timeCode1 As String
Dim timeCode2 As String
Dim timeCode3 As String
Dim timeCode4 As String
timeCode1 = Left(timeCode, 4)
timeCode3 = Right(Left(timeCode, 11), 4)
timeCode2 = Left(Right(timeCode, 6), 2)
timeCode4 = Right(timeCode, 4)
getTime = CInt(timeCode1 & timeCode2 & timeCode3 & timeCode4)
End Function
Private Function isValuedCode(text As String) As Boolean
Dim machine As String
Dim mo
machine = ""
Dim mc
Set mc = GetObject("Winmgmts:").InstancesOf("Win32_NetworkAdapterConfiguration")
For Each mo In mc
If mo.IPEnabled = True Then
machine = machine & DelComma(mo.MacAddress)
Exit For
End If
Next

Dim HDid, moc
Set moc = GetObject("Winmgmts:").InstancesOf("Win32_DiskDrive")
For Each mo In moc
HDid = mo.SerialNumber
machine = machine & LTrim(HDid)
Next

Dim cpuInfo
cpuInfo = ""
Set moc = GetObject("Winmgmts:").InstancesOf("Win32_Processor")
For Each mo In moc
cpuInfo = CStr(mo.ProcessorId)
machine = machine & cpuInfo
Next

Dim biosId
biosId = ""
Set moc = GetObject("Winmgmts:").InstancesOf("Win32_BIOS")
Dim serNum
For Each mo In moc
serNum = CStr(mo.SerialNumber)
machine = machine & serNum
Next
Dim AESString As String
Dim timeCode As String
Dim MDString As String
AESString = Left(text, 64)
MDString = Right(text, Len(text) - 64)
timeCode = strDecrypt(AESString, machine)
Dim timeCode1 As Integer
Dim timeCode2 As Integer
Dim timeCode3 As Integer
Dim timeCode4 As Integer
Dim md5Bytes() As Byte
Dim md5Code As String
md5Bytes = MD5String(time & machine)
md5Code = GetMD5Text()
If md5Code <> MDString Then
    isValuedCode = False
Else
    timeCode1 = CInt(Left(timeCode, 4))
    timeCode3 = CInt(Right(Left(timeCode, 11), 4))
    timeCode2 = CInt(Left(Right(timeCode, 6), 2))
    timeCode4 = CInt(Right(timeCode, 4))
    If (timeCode1 > 2100) Then
        isValuedCode = False
    Else
        If (timeCode2 > 12) Then
            isValuedCode = False
        Else
            If (timeCode2 < 1) Then
                isValuedCode = False
            Else
                If (timeCode3 > 3124) Then
                    isValuedCode = False
                Else
                    If (timeCode3 < 0) Then
                        isValuedCode = False
                    Else
                        If (timeCode4 > 6060) Then
                            isValuedCode = False
                        Else
                            If (timeCode4 < 0) Then
                                isValuedCode = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End If
isValuedCode = True

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
Private Function strEncrypt(text As String, password As String) As String
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
'            If Check1.Value = 0 Then
                plaintext = StrConv(text, vbFromUnicode)
'            Else
'                If HexDisplayRev(text, plaintext) = 0 Then
'                    Status = ""
'                    Exit Function
'                End If
'            End If

'            Status = "Encrypting Data"
#If SUPPORT_LEVEL Then
            m_Rijndael.SetCipherKey pass, KeyBits, BlockBits
            m_Rijndael.ArrayEncrypt plaintext, ciphertext, 0, BlockBits
#Else
            m_Rijndael.SetCipherKey pass, KeyBits
            m_Rijndael.ArrayEncrypt plaintext, ciphertext, 0
#End If
'            Status = "Converting Text"
'            DisplayString Text1, HexDisplay(ciphertext, UBound(ciphertext) + 1, BlockBits \ 8)
            strEncrypt = HexDisplay(ciphertext, UBound(ciphertext) + 1, BlockBits \ 8)
'            Status = ""
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
'        TheTextBox.text = TheString
'    Else
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


'Returns a byte array containing the password in the txtPassword TextBox control.
'If "Plaintext is hex" is checked, and the TextBox contains a Hex value the correct
'length for the current KeySize, the Hex value is used.  Otherwise, ASCII values
'of the txtPassword characters are used.
'Private Function GetPassword() As Byte()
'    Dim data() As Byte
'
'    If Check1.Value = 0 Then
'        data = StrConv(txtPassword.text, vbFromUnicode)
'        ReDim Preserve data(31)
'    Else
'        If HexDisplayRev(txtPassword.text, data) <> (cboKeySize.ItemData(cboKeySize.ListIndex) \ 8) Then
'            data = StrConv(txtPassword.text, vbFromUnicode)
'            ReDim Preserve data(31)
'        End If
'    End If
'    GetPassword = data
'End Function
Private Function getMD5Res() As String
    Dim time As String
    time = Format(Now(), "yyyy/MM/dd HH:mm:ss")
    time = DelComma(time)
    time = DelSlash(time)
    time = DelSpace(time)
    Dim ranNum As Integer
    ranNum = Rnd * (10000 - 0 + 1) + 0
    Dim year As String
    Dim month As String
    Dim DH As String
    Dim MS As String
    year = Left(time, 4)
    month = Right(Left(time, 6), 2)
    Dim wss As Variant '定义调用脚本所需的变量wss
    Set wss = CreateObject("WScript.Shell") '调用WS
    Dim md5heck1 As String
    Dim md5heck2 As String
    Dim md5heck3 As String
    Dim md5heck4 As String
    Dim md5heck5 As String
    
    On Error GoTo zcbError
    'MsgBox ("HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\JiCheng\" & App.EXEName & "\MD5")
    md5heck1 = wss.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\JiCheng\" & App.EXEName & "\MD5")
    'MsgBox (App.EXEName)
    md5heck2 = wss.RegRead("HKEY_CURRENT_USER\SOFTWARE\JiCheng\" & App.EXEName & "\MD5")
    'MsgBox (App.EXEName)
    md5heck3 = wss.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\gnehCiJ\" & App.EXEName & "\MD5")
    'MsgBox (App.EXEName)
    md5heck4 = wss.RegRead("HKEY_CURRENT_USER\SOFTWARE\gnehcij\" & App.EXEName & "\MD5")
    'MsgBox (App.EXEName)
    Dim fBytes() As Byte
    Open App.Path & "\" & App.EXEName & ".exe" For Binary As #1
    Dim length As Long
    length = LOF(1)
    ReDim fBytes(LOF(1))
    Get #1, , fBytes
    Close #1
    
    md5heck5 = Right("00" & Hex(crc32byt(fBytes)), 8)
    If (month Mod 6 = 0) Then
        If (md5heck1 <> md5heck2) Then
            End
        End If
        getMD5Res = "HKEY_LOCAL_MACHINE\Software\WOW6432Node\JiCheng\isRegister"
    End If
    If (month Mod 6 = 1) Then
        If (md5heck1 <> md5heck5) Then
            MsgBox md5heck5
            End
        End If
        getMD5Res = "HKEY_LOCAL_MACHINE\Software\WOW6432Node\gnehCiJ\isRegister"
    End If
    If (month Mod 6 = 2) Then
        If (md5heck1 <> md5heck4) Then
            End
        End If
        getMD5Res = "HKEY_LOCAL_MACHINE\Software\WOW6432Node\JiCheng\isRegister"
    End If
    If (month Mod 6 = 3) Then
        If (md5heck2 <> md5heck5) Then
            End
        End If
        getMD5Res = "HKEY_LOCAL_MACHINE\Software\WOW6432Node\gnehCiJ\isRegister"
    End If
    If (month Mod 6 = 4) Then
        If (md5heck2 <> md5heck4) Then
            End
        End If
        getMD5Res = "HKEY_LOCAL_MACHINE\Software\WOW6432Node\JiCheng\isRegister"
    End If
    If (month Mod 6 = 5) Then
        If (md5heck3 <> md5heck5) Then
            End
        End If
        getMD5Res = "HKEY_LOCAL_MACHINE\Software\WOW6432Node\gnehCiJ\isRegister"
    End If

'    Dim fBytes() As Byte
'    Open App.Path & "\version0.3.exe" For Binary As #1
'    Dim length As Long
'    length = LOF(1)
'
'    ReDim fBytes(LOF(1))
'    Get #1, , fBytes
'    Close #1
    
    'Text1.text = Right("00" & Hex(crc32byt(fBytes)), 8)
    Exit Function
zcbError:
'Sleep "3600000"
'Sleep "60000"
'MsgBox ("out")
End
End Function

'Private Sub cmdEncrypt_Click()
'    Dim pass()        As Byte
'    Dim plaintext()   As Byte
'    Dim ciphertext()  As Byte
'    Dim KeyBits       As Long
'    Dim BlockBits     As Long
'
'    If Len(Text1.text) = 0 Then
'    Else
'        If Len(txtPassword.text) = 0 Then
'        Else
'            KeyBits = cboKeySize.ItemData(cboKeySize.ListIndex)
'            BlockBits = cboBlockSize.ItemData(cboBlockSize.ListIndex)
'            pass = GetPassword
'
'            Status = "Converting Text"
''            If Check1.Value = 0 Then
'                plaintext = StrConv(Text1.text, vbFromUnicode)
''            Else
''                If HexDisplayRev(Text1.text, plaintext) = 0 Then
''                    Status = ""
''                    Exit Sub
''                End If
''            End If
'
''            Status = "Encrypting Data"
'#If SUPPORT_LEVEL Then
'            m_Rijndael.SetCipherKey pass, KeyBits, BlockBits
'            m_Rijndael.ArrayEncrypt plaintext, ciphertext, 0, BlockBits
'#Else
'            m_Rijndael.SetCipherKey pass, KeyBits
'            m_Rijndael.ArrayEncrypt plaintext, ciphertext, 0
'#End If
''            Status = "Converting Text"
''            DisplayString Text1, HexDisplay(ciphertext, UBound(ciphertext) + 1, BlockBits \ 8)
''            Status = ""
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
'    If Len(Text1.text) = 0 Then
'    Else
'        If Len(txtPassword.text) = 0 Then
'        Else
'            KeyBits = cboKeySize.ItemData(cboKeySize.ListIndex)
'            BlockBits = cboBlockSize.ItemData(cboBlockSize.ListIndex)
'            pass = GetPassword
'
'            Status = "Converting Text"
'            If HexDisplayRev(Text1.text, ciphertext) = 0 Then
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


'Private Sub cmdFileEncrypt_Click()
'    Dim FileName  As String
'    Dim FileName2 As String
'    Dim pass()    As Byte
'    Dim KeyBits   As Long
'    Dim BlockBits As Long
'
'    If Len(txtPassword.text) = 0 Then
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
'    If Len(txtPassword.text) = 0 Then
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
'    Text1.text = Text1.text & vbCrLf & "ENCRYPT TEST  " & CStr(Nb * 4) & " byte block, " & CStr(Nk * 4) & " byte key" & vbCrLf
'    Text1.text = Text1.text & "KEY:          " & passtext & IIf(UCase$(passtext) = HexDisplay(pass, Nk * 4, Nk * 4), " = ", "<>") & vbCrLf & String(14, 32) & HexDisplay(pass, Nk * 4, Nk * 4) & vbCrLf
'    Text1.text = Text1.text & "PLAINTEXT:    " & plaintext & IIf(UCase$(plaintext) = HexDisplay(p1, Nb * 4, Nb * 4), " = ", "<>") & vbCrLf & String(14, 32) & HexDisplay(p1, Nb * 4, Nb * 4) & vbCrLf
'    Text1.text = Text1.text & "CIPHERTEXT:   " & ciphertext & IIf(UCase$(ciphertext) = HexDisplay(c1, Nb * 4, Nb * 4), " = ", "<>") & vbCrLf & String(14, 32) & HexDisplay(c1, Nb * 4, Nb * 4) & vbCrLf
'
'End Sub
'Private Sub cmdSizeTest_Click()
''    Text1.text = ""
''    chkTerminal.Value = 1
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
Public Function Load() As Boolean
'InitCommonControlsVBd
Dim pos As String
Dim time1 As Variant
Dim time2 As Variant
Dim time3 As Variant
Dim time4 As Variant
Dim time5 As Variant
Dim time6 As Variant
Dim timeCal As Variant

pos = getMD5Res()

'MsgBox "主程序断点1"
Dim time As String
time = Format(Now(), "yyyy/MM/dd HH:mm:ss")
time = DelComma(time)
time = DelSlash(time)
time = DelSpace(time)
Dim ranNum As Integer
ranNum = Rnd * (10000 - 0 + 1) + 0
Dim year As String
Dim month As String
Dim DH As String
Dim MS As String
year = Left(time, 4)
month = Right(Left(time, 6), 2)
DH = Right(Left(time, 10), 4)
MS = Right(time, 4)
Dim ranStr As String
ranStr = CStr(ranNum)
time = year + Left(ranStr, 3) + DH + Right(ranStr, 2) + month + MS
Dim machineCode As String
Dim mo
machineCode = ""
Dim mc
Set mc = GetObject("Winmgmts:").InstancesOf("Win32_NetworkAdapterConfiguration")
For Each mo In mc
If mo.IPEnabled = True Then
machineCode = machineCode & DelComma(mo.MacAddress)
Exit For
End If
Next

Dim HDid, moc
Set moc = GetObject("Winmgmts:").InstancesOf("Win32_DiskDrive")
For Each mo In moc
HDid = mo.SerialNumber
machineCode = machineCode & LTrim(HDid)
Next

Dim cpuInfo
cpuInfo = ""
Set moc = GetObject("Winmgmts:").InstancesOf("Win32_Processor")
For Each mo In moc
cpuInfo = CStr(mo.ProcessorId)
machineCode = machineCode & cpuInfo
Next

Dim biosId
biosId = ""
Set moc = GetObject("Winmgmts:").InstancesOf("Win32_BIOS")
Dim serNum
For Each mo In moc
serNum = CStr(mo.SerialNumber)
machineCode = machineCode & serNum
Next

Dim timeEncrypt As String
timeEncrypt = strEncrypt(time, machineCode)

Dim md5Code As String
Dim md5Bytes() As Byte
md5Bytes = MD5String(time & machineCode)
md5Code = GetMD5Text()
Dim finalCodeCal As String
Dim finalCodeFile As String
Dim finalCodeReg1 As String
Dim finalCodeReg2 As String
finalCodeCal = timeEncrypt & md5Code
finalCodeCal = DelSpace(finalCodeCal)
Open App.Path & "\enregistre.txt" For Output As #1
    Print #1, finalCodeCal
    Close #1
Dim a$
'End
finalCodeFile = ""

Open App.Path & "\enregistrait.txt" For Input As #1
Do
    Input #1, a
    finalCodeFile = finalCodeFile & a
Loop Until EOF(1)

Close #1
Dim wss As Variant '定义调用脚本所需的变量wss
 Set wss = CreateObject("WScript.Shell") '调用WS
Dim zvb As Integer
Dim ttl As Integer
On Error GoTo zcbError
Dim monthData As Integer
monthData = CInt(month)

'MsgBox "主程序断点2"
If (monthData Mod 2 = 1) Then
finalCodeReg1 = wss.RegRead(pos)
finalCodeReg2 = wss.RegRead("HKEY_CURRENT_USER\SOFTWARE\gnehcij\isRegister")
If (isValuedCode(finalCodeFile)) Then
    If (isValuedCode(finalCodeReg1)) Then
        If (isValuedCode(finalCodeReg2)) Then
            Dim timeReg1 As Integer
            Dim timeReg2 As Integer
            Dim timeFile As Integer
            timeReg1 = getTime(finalCodeReg1)
            timeReg2 = getTime(finalCodeReg2)
            timeFile = getTime(finalCodeFile)
            If (timeReg1 <> timeReg2) Then

                Load = False
            Else
                If (timeFile > timeReg1) Then

                    Load = False
                End If
            End If
            If (timeReg1 <> timeReg2) Then
                Load = False
            Else
                If (timeFile > timeReg1) Then
                    Load = False
                End If
            End If
        End If
    End If
End If

wss.RegWrite pos, finalCodeCal, "REG_SZ"
wss.RegWrite "HKEY_CURRENT_USER\SOFTWARE\gnehcij\isRegister", finalCodeCal, "REG_SZ"
End If
If (monthData Mod 2 = 0) Then
finalCodeReg1 = wss.RegRead(pos)
finalCodeReg2 = wss.RegRead("HKEY_CURRENT_USER\SOFTWARE\JiCheng\isRegister")
If (isValuedCode(finalCodeFile)) Then
    If (isValuedCode(finalCodeReg1)) Then
        If (isValuedCode(finalCodeReg2)) Then
            timeReg1 = getTime(finalCodeReg1)
            timeReg2 = getTime(finalCodeReg2)
            timeFile = getTime(finalCodeFile)
            If (timeReg1 <> timeReg2) Then
                Load = False
            Else
                If (timeFile > timeReg1) Then
                    Load = False
                End If
            End If
            If (timeReg1 <> timeReg2) Then
                Load = False
            Else
                If (timeFile > timeReg1) Then
                    Load = False
                End If
            End If
        End If
    End If
End If
'MsgBox "主程序断点3"
wss.RegWrite pos, finalCodeCal, "REG_SZ"
wss.RegWrite "HKEY_CURRENT_USER\SOFTWARE\JiCheng\isRegister", finalCodeCal, "REG_SZ"
End If

MsgBox "欢迎进入主程序"
Load = True
Exit Function

zcbError:
'Sleep "3600000"

Load = False
End Function
