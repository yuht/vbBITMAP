VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6015
   ClientLeft      =   6960
   ClientTop       =   3600
   ClientWidth     =   8670
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6015
   ScaleMode       =   0  'User
   ScaleWidth      =   8670
   Begin VB.TextBox Text5 
      Height          =   1275
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "Form1.frx":030A
      Top             =   4680
      Width           =   6540
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   7515
      Top             =   270
   End
   Begin VB.TextBox Text4 
      Height          =   1215
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "Form1.frx":0310
      Top             =   3375
      Width           =   6495
   End
   Begin VB.TextBox Text3 
      Height          =   1215
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Form1.frx":0316
      Top             =   2070
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
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form1.frx":031C
      Top             =   765
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
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
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      DrawMode        =   3  'Not Merge Pen
      Height          =   135
      Index           =   999
      Left            =   6720
      Shape           =   1  'Square
      Top             =   1380
      Visible         =   0   'False
      Width           =   135
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
  Me.Command1.Enabled = False
  SaveChinese
  Chinese2HData
  PrintHData
  H2V
  PrintV
  H2V2
  PrintV2
  Me.Cls
End Sub



Private Sub Form_Load()
  '**************************************
  ZitiPath = App.Path & "\" & "hzk16" '
  Text2.Text = "横向取反"
  Text3.Text = "横向计算结果"
  Command1.Caption = "生成字模"
  E = 1
  Text4.Text = "竖向计算结果" & vbCrLf & vbCrLf & "点阵上用的是这个数据"
  
  '=================================
  'Shape1 组成的16*16点阵
  On Error Resume Next
  Dim I, J, SI As Integer
  For I = 0 To 15 '16行
    For J = 0 To 15 '16列
      SI = I * 16 + J   '每行16个
      Load Shape1(SI)
      Shape1(SI).Top = Shape1(SI - J - 1).Top + 120
      If J = 0 Then
        Shape1(SI).Left = Shape1(0).Left
      Else
        Shape1(SI).Left = Shape1(SI - 1).Left + 120
      End If
      'Shape1(SI).Visible = True
    Next
  Next
  Command1_Click
End Sub

'=======================
'从Text1中找出汉字,存放到HzStr1数组中
'汉字个数存放到 N 中
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

'========================
' 汉字顺序横排取模,存放到ZMHP数组中
' 读字模文件存放到Hzk166中
' 找到汉字字模位置,并按行顺序存放到ZMHP(汉字个数,32)数组中
'
Function Chinese2HData()
  Dim Hzk166() As Byte '字库数组
  Dim Sum, QWM, QM, WM  '文件长度,区位码,区码,位码
  Dim intFileNum  '文件号
  Dim FileNa  '文件名
  Dim I, J As Integer
  
'===============
'读取字库到数组
  
  FileNa = ZitiPath '字模路径
  intFileNum = FreeFile '空闲文件号
  Open FileNa For Binary As #intFileNum   '打开文件
  Sum = LOF(intFileNum)   '文件长度
  ReDim Hzk166(1 To Sum)  '按文件长度定义数组
  Get #intFileNum, , Hzk166 '存入到数组 Hzk166
  Close #intFileNum '关闭字库文件,防止发生错误
  
  ReDim ZMHP(1 To N, 1 To 32) '定义数组 N维

  '===================
  '获取汉字区位码
  For I = 1 To UBound(HzStr1)
    QWM = Hex(Asc(HzStr1(I)) - &HA0A0) '区位码
    If Len(QWM) = 3 Then
      QM = Left(QWM, 1) '区码
    ElseIf Len(QWM) = 4 Then
      QM = Left(QWM, 2) '区码
    End If
    WM = Right(QWM, 2) '位码

    '================
    '获取汉字点阵起始位置
    Address1 = 32 * ((CLng("&H" & QM) - 1) * 94 + (CLng("&H" & WM) - 1))
    For J = 1 To 32 '每个字为32个字节
      ZMHP(I, J) = Hzk166(Address1 + J)    '将点阵数据存入,数组
    Next
  Next
End Function
  
'=============================
'输出横排取模的数组ZMHP数据,并用Shape1 显示
'
'
Function PrintHData()
  Dim S, W As String '存放原码,取反吗
  Dim I, J As Integer 'for计数器
  Dim tmpstr As String  '临时字符 存放zmhp(j)
  

  For J = 1 To 32
    tmpstr = Hex(ZMHP(1, J))
    If Len(tmpstr) = 1 Then tmpstr = "0" & tmpstr '格式化该字符,长度不为2用0补.
    '==================左侧字符表显示该字符
        
    '显示shape1
    For I = 1 To 8
      '延时Timer1
'      Timer1.Enabled = True
'      While Timer1.Enabled = True
'      DoEvents
'      Wend
      '根据相应位控制相对应Shape1
      If (Val("&h" & tmpstr) And 2 ^ (8 - I)) = 0 Then
        Shape1((J - 1) * 8 + I - 1).Visible = True
      Else
        Shape1((J - 1) * 8 + I - 1).Visible = False
      End If
    Next


    S = S & "0x" & tmpstr & ","

    tmpstr = Hex(&HFF - ZMHP(1, J))
    If Len(tmpstr) = 1 Then tmpstr = "0" & tmpstr
    W = W & "0x" & tmpstr & ","

    If J Mod 8 = 0 Then
      S = S & vbCrLf
      W = W & vbCrLf
    End If
  Next

  Text2 = "横向取原数据：" & vbCrLf & S
  Text3 = "横向取反后数据：" & vbCrLf & W

End Function

'============================================
'将ZMHP 横排字符数组生成数组ZMSP顺序纵向/竖排数组ZMSP
'将ZMHP的1,3,5,7~~~~ 15 的第8位作为ZMSP的第1个
'将ZMHP的17,16,21,23~~~~ 32 的第8位作为ZMSP的第2个
'
'返回结果存放到ZMSP
'
Function H2V() '将横排转化为竖排
  Dim I, K, J, M As Integer
  Dim BitHZ1 As Byte '用来判断该位的值
  Dim Z, ZA As Byte
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
        Z = 0
        ZA = 0
        
        If ((ZMHP(I, K) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '8
        Z = &H80 * BitHZ1 '作位最高位
        If ((ZMHP(I, K + 16) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '8
        ZA = &H80 * BitHZ1
        
        
        If ((ZMHP(I, K + 2) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '7
        Z = Z + (&H40 * BitHZ1)
        If ((ZMHP(I, K + 18) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '7
        ZA = ZA + (&H40 * BitHZ1)
        
        
        If ((ZMHP(I, K + 4) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '6
        Z = Z + (&H20 * BitHZ1)
        If ((ZMHP(I, K + 20) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '6
        ZA = ZA + (&H20 * BitHZ1)
        
        
        If ((ZMHP(I, K + 6) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '5
        Z = Z + (&H10 * BitHZ1)
        If ((ZMHP(I, K + 22) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '5
        ZA = ZA + (&H10 * BitHZ1)
        
        If ((ZMHP(I, K + 8) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '4
        Z = Z + (&H8 * BitHZ1)
        If ((ZMHP(I, K + 24) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '4
        ZA = ZA + (&H8 * BitHZ1)
        
        If ((ZMHP(I, K + 10) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '3
        Z = Z + (&H4 * BitHZ1)
        If ((ZMHP(I, K + 26) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '3
        ZA = ZA + (&H4 * BitHZ1)
        
        If ((ZMHP(I, K + 12) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '2
        Z = Z + (&H2 * BitHZ1)
        If ((ZMHP(I, K + 28) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '2
        ZA = ZA + (&H2 * BitHZ1)
        
        If ((ZMHP(I, K + 14) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '1
        Z = Z + (&H1 * BitHZ1) '作为最低位
        If ((ZMHP(I, K + 30) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '1
        ZA = ZA + (&H1 * BitHZ1)
        
        ZMSP(I, J) = Z '取的为上部分
        J = J + 1
        
        ZMSP(I, J) = ZA '下一部分
        J = J + 1 '            这一部分有些麻烦,要看的话多注意理解vb中16进制数的运算

        
        BitI = (BitI / &H2) '取横排的下一位*****************哈哈，但愿大家能看懂××××××××××××××
      Next
    Next
  Next
End Function '将横排存放的汉字转化位,竖排存放的由于vb没有移位运算只有按位相与,进行判断.

Function PrintV()
  Dim I  As Integer
  Dim S, W As String
  Dim tmpstr As String
  
  For I = 1 To 32
  tmpstr = Hex(ZMSP(1, I))
  If Len(tmpstr) = 1 Then tmpstr = "0" & tmpstr '格式化,长度变为2 不足前面用0补
  'Debug.Print Hex(&HFF - zmhp(1, i)) & ",";
  S = S & "0x" & tmpstr & ","
  If I Mod 8 = 0 Then S = S & vbCrLf
  Next


  Text4.Text = "竖向取原数据：" & vbCrLf & S
  'Text2.Text = "横向取反后数据：" & vbCrLf & w

End Function

Function H2V2() '将横排转化为竖排
  Dim I, K, J, M As Integer
  Dim BitHZ1 As Byte '用来判断该位的值
  Dim Z, ZA As Byte
  Dim BitI As Byte  '横向取模转纵向取模的第I位, I= 8~1
  ReDim ZMSP(1 To N, 1 To 32)
  Dim tmpM22 As Byte
  Dim M2 As Integer
  For I = 1 To N
    'i = 1
    J = 1
    For K = 1 To 2 '
      'Debug.Print "_____"
      BitI = &H80 '
      'If (biti >= &H1) Then
      For M = 1 To 8
        'Debug.Print "_____"'运算根据16*16点阵横排与竖排的存放特性进行计算.
        Z = 0
        ZA = 0
          tmpM22 = &H80
        For M2 = 0 To 7
        
          If ((ZMHP(I, K + M2 * 2) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '8
          Z = Z + tmpM22 * BitHZ1 '作位最高位
          If ((ZMHP(I, K + M2 * 2 + 16) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '8
          ZA = ZA + tmpM22 * BitHZ1
                
          tmpM22 = tmpM22 / &H2
        Next
        BitI = BitI / &H2
        
        ZMSP(I, J) = Z '取的为上部分
        Debug.Print J
        ZMSP(I, J + 1) = ZA '下一部分
        
        J = J + 2 '            这一部分有些麻烦,要看的话多注意理解vb中16进制数的运算


      Next
    Next
  Next
End Function '将横排存放的汉字转化位,竖排存放的由于vb没有移位运算只有按位相与,进行判断.

Function PrintV2()
  Dim I  As Integer
  Dim S, W As String
  Dim tmpstr As String
  
  For I = 1 To 32
  tmpstr = Hex(ZMSP(1, I))
  If Len(tmpstr) = 1 Then tmpstr = "0" & tmpstr '格式化,长度变为2 不足前面用0补
  'Debug.Print Hex(&HFF - zmhp(1, i)) & ",";
  S = S & "0x" & tmpstr & ","
  If I Mod 8 = 0 Then S = S & vbCrLf
  Next


  Text5.Text = "竖向取原数据2：" & vbCrLf & S
  'Text2.Text = "横向取反后数据：" & vbCrLf & w

End Function


Private Sub Timer1_Timer()
  Timer1.Enabled = False
End Sub

'程序设计李健-2007-12-16 QQ:102126913 tel:
'write for 穿透瞬间
'*******************************************************
'rewrite 2008-1-8
'write for snail boy

'REWRITE 2009-05-16
'WRITE BY LOVEYU
'WRITE FOR MYSELF


