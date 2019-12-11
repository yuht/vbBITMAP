VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4710
   ClientLeft      =   6975
   ClientTop       =   3615
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleMode       =   0  'User
   ScaleWidth      =   6660
   Begin VB.TextBox Text4 
      Height          =   1215
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "Form1.frx":0000
      Top             =   3420
      Width           =   6495
   End
   Begin VB.TextBox Text3 
      Height          =   1215
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Form1.frx":0006
      Top             =   780
      Width           =   6495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   3900
      TabIndex        =   2
      Top             =   60
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   1215
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form1.frx":000C
      Top             =   2100
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      TabIndex        =   0
      Text            =   "于"
      Top             =   60
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim HzStr1() As String '从文本中取出的中文存放的一个数组
Dim ZitiPath   As String '字体所在路径

Dim N As Integer '文本计数用hzLen()
Dim Address1  '存放在hzk16中的地址
Dim ZMHP(), ZMSP() As Byte '定义横/竖排字存放的树组
Dim E
 

Private Sub Command1_Click()
  SaveChinese
  Chinese2Data
  Data2H
  H2V
  PrintV
  Me.Cls
End Sub



Private Sub Form_Load()
  '**************************************
  ZitiPath = App.Path & "\" & "hzk16" '
  Text2.Text = "横向取反"
  'Me.Caption = "write for:——————"
  Text3.Text = "横向计算结果"
  Command1.Caption = "生成字模"
  E = 1
  Text4.Text = "竖向计算结果" & vbCrLf & vbCrLf & "点阵上用的是这个数据"
End Sub


Function SaveChinese()

  ReDim HzStr1(0)
  
  Dim UboundStr1, I
  For I = 1 To Len(Text1)
    If Asc(Mid(Text1.Text, I, 1)) < 0 Then
      UboundStr1 = UBound(HzStr1) + 1
      ReDim Preserve HzStr1(UboundStr1)
      HzStr1(UboundStr1) = Mid(Text1.Text, I, 1) '把汉字存入数组中.
    End If
  Next
  
  N = UboundStr1

End Function

Function Chinese2Data()
  Dim Hzk166() As Byte '字库数组
  Dim Sum, QWM, QM, WM  '文件长度,区位码,区码,位码
  Dim I, J As Integer
  Dim intFileNum  '文件号
  Dim FileNa  '文件名
'===============
'读取字库到数组
  
  FileNa = ZitiPath
  intFileNum = FreeFile
  Open FileNa For Binary As #intFileNum   '打开文件
  Sum = LOF(intFileNum)   '文件长度
  ReDim Hzk166(1 To Sum)  '按文件长度定义数组
  Get #intFileNum, , Hzk166 '存入数组
  Close #intFileNum '关闭字库文件,防止发生错误
  
  ReDim ZMHP(1 To N, 1 To 32) 'As Byte

  '===================
  '获取汉字区位码
  For I = 1 To UBound(HzStr1)
    QWM = Hex(Asc(HzStr1(I)) - &HA0A0)
    If Len(QWM) = 3 Then
      QM = Left(QWM, 1)
    ElseIf Len(QWM) = 4 Then
      QM = Left(QWM, 2)
    End If
    WM = Right(QWM, 2)

    '================
    '获取汉字点阵
    Address1 = 32 * ((CLng("&H" & QM) - 1) * 94 + (CLng("&H" & WM) - 1))
    For J = 1 To 32 '每个字为32个字节
      ZMHP(I, J) = Hzk166(Address1 + J)    '将点阵数据存入,数组
'      On Error Resume Next
    Next
  Next
  End Function
  
Function Data2H()
  Dim S, W As String
  Dim I, J As Integer
  Dim StartN, EndN As Integer
  Dim TmpStr As String
  For I = 1 To 4
    EndN = I * 8
    StartN = EndN - 7
    
    For J = StartN To EndN
      TmpStr = Hex(ZMHP(1, J))
      If Len(TmpStr) = 1 Then TmpStr = "0" & TmpStr
      S = S & "0x" & TmpStr & ","
      
      TmpStr = Hex(&HFF - ZMHP(1, J))
      If Len(TmpStr) = 1 Then TmpStr = "0" & TmpStr
      W = W & "0x" & TmpStr & ","
    Next
    S = S & vbCrLf
    W = W & vbCrLf
  Next

  Text3 = "横向取原数据：" & vbCrLf & S
  Text2 = "横向取反后数据：" & vbCrLf & W

End Function

'============================================
'将ZMHP 横排字符数组转为竖排数组ZMSP
'将ZMHP的1,3,5,7~~~~ 15 的第8位作为ZMSP的第1个
'将ZMHP的17,16,21,23~~~~ 32 的第8位作为ZMSP的第2个
'
'返回结果存放到ZMSP
'
Function H2V() '将横排转化为竖排
  Dim I, K, J, M As Integer
  Dim BitHZ1 As Byte '用来判断该位的值
  Dim Z As Byte
  Dim BitI As Byte  '横向取模转纵向取模的第I位, I= 8~1
  ReDim ZMSP(1 To N, 1 To 32)
  
  For I = 1 To N
    'i = 1
    J = 1
    For K = 1 To 2 '
      'Debug.Print "_____"
      BitI = &H80 '
      'If (biti >= &H1) Then
      For M = 1 To 8
        'Debug.Print "_____"'运算根据16*16点阵横排与竖排的存放特性进行计算.
        Z = &H0
        
        If ((ZMHP(I, K) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '8
        Z = &H80 * BitHZ1 '作位最高位
        
        If ((ZMHP(I, K + 2) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '7
        Z = Z + (&H40 * BitHZ1)
        
        If ((ZMHP(I, K + 4) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '6
        Z = Z + (&H20 * BitHZ1)
        
        If ((ZMHP(I, K + 6) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '5
        Z = Z + (&H10 * BitHZ1)
        
        If ((ZMHP(I, K + 8) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '4
        Z = Z + (&H8 * BitHZ1)
        
        If ((ZMHP(I, K + 10) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '3
        Z = Z + (&H4 * BitHZ1)
        
        If ((ZMHP(I, K + 12) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '2
        Z = Z + (&H2 * BitHZ1)
        
        If ((ZMHP(I, K + 14) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '1
        Z = Z + (&H1 * BitHZ1) '作为最低位
        
        ZMSP(I, J) = Z '取的为上部分
        J = J + 1
        Z = 0
        
        If ((ZMHP(I, K + 16) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '8
        Z = &H80 * BitHZ1
        
        If ((ZMHP(I, K + 18) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '7
        Z = Z + (&H40 * BitHZ1)
        
        If ((ZMHP(I, K + 20) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '6
        Z = Z + (&H20 * BitHZ1)
        
        If ((ZMHP(I, K + 22) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '5
        Z = Z + (&H10 * BitHZ1)
        
        If ((ZMHP(I, K + 24) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '4
        Z = Z + (&H8 * BitHZ1)
        
        If ((ZMHP(I, K + 26) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '3
        Z = Z + (&H4 * BitHZ1)
        
        If ((ZMHP(I, K + 28) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '2
        Z = Z + (&H2 * BitHZ1)
        
        If ((ZMHP(I, K + 30) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '1
        Z = Z + (&H1 * BitHZ1)
        
        ZMSP(I, J) = Z '下一部分
        J = J + 1 '            这一部分有些麻烦,要看的话多注意理解vb中16进制数的运算
        Z = 0
        
        BitI = (BitI / &H2) '取横排的下一位*****************哈哈，但愿大家能看懂××××××××××××××
      Next
    Next
  Next
End Function '将横排存放的汉字转化位,竖排存放的由于vb没有移位运算只有按位相与,进行判断.

Function PrintV()
  Dim I  As Integer
  Dim S, W As String
  For I = 1 To 8
  'Debug.Print Hex(&HFF - zmhp(1, i)) & ",";
  S = S & "0x" & Hex((ZMSP(1, I))) & ","
  Next
  
  S = S & vbCrLf
  
  For I = 9 To 16
  'Debug.Print Hex(&HFF - zmhp(1, i)) & ",";
  S = S & "0x" & Hex((ZMSP(1, I))) & ","
  Next
  
  S = S & vbCrLf
  
  For I = 17 To 24
  'Debug.Print Hex(&HFF - zmhp(1, i)) & ",";
  S = S & "0x" & Hex((ZMSP(1, I))) & ","
  Next I
  
  S = S & vbCrLf
  
  For I = 25 To 32
  'Debug.Print Hex(&HFF - zmhp(1, i)) & ",";
  S = S & "0x" & Hex((ZMSP(1, I))) & ","
  Next I
  
  S = S & vbCrLf

  Text4.Text = "竖向取原数据：" & vbCrLf & S
  'Text2.Text = "横向取反后数据：" & vbCrLf & w

End Function

'程序设计李健-2007-12-16 QQ:102126913 tel:
'write for 穿透瞬间
'*******************************************************
'rewrite 2008-1-8
'write for snail boy
