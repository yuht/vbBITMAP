VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   10530
   ClientLeft      =   6960
   ClientTop       =   3600
   ClientWidth     =   18630
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10530
   ScaleMode       =   0  'User
   ScaleWidth      =   18630
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      Caption         =   "字体预览"
      Height          =   2670
      Left            =   8730
      TabIndex        =   0
      Top             =   1215
      Width           =   7800
      Begin VB.HScrollBar HScroll1 
         Height          =   240
         Left            =   450
         MousePointer    =   9  'Size W E
         TabIndex        =   11
         Top             =   2295
         Visible         =   0   'False
         Width           =   6765
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000009&
         BackStyle       =   1  'Opaque
         DrawMode        =   3  'Not Merge Pen
         Height          =   135
         Index           =   999
         Left            =   495
         Shape           =   1  'Square
         Top             =   270
         Visible         =   0   'False
         Width           =   135
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "字模输出"
      Height          =   7890
      Left            =   45
      TabIndex        =   4
      Top             =   1215
      Width           =   8610
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   3420
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Text            =   "Form1.frx":030A
         Top             =   240
         Width           =   8430
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   4080
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Text            =   "Form1.frx":0310
         Top             =   3720
         Width           =   8430
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "输入字符"
      Height          =   1050
      Left            =   45
      TabIndex        =   1
      Top             =   90
      Width           =   16035
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   180
         TabIndex        =   3
         Text            =   "于海涛啊哇哈嘿咦"
         Top             =   240
         Width           =   14175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   615
         Left            =   14490
         TabIndex        =   2
         Top             =   225
         Width           =   1455
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   135
      Top             =   90
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   900
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Text            =   "Form1.frx":0316
      Top             =   5265
      Width           =   8430
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Height          =   900
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Text            =   "Form1.frx":031E
      Top             =   6210
      Width           =   8430
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      Height          =   900
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Text            =   "Form1.frx":0326
      Top             =   7155
      Width           =   8430
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      Height          =   900
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Text            =   "Form1.frx":032E
      Top             =   8100
      Width           =   8430
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
Dim ZMHP(), ZMSP(), ZMVH() As Byte '定义横/竖排字存放的树组
Dim E
 

Private Sub Command1_Click()
  Me.Cls
'  Me.Command1.Enabled = False
  SaveChinese '从 text1中找出汉字
  
  If N = 0 Then
    'MsgBox "没有汉字!", vbInformation + vbOKOnly, "提示!"
    Text1 = "于海涛"
    Text1.SetFocus
    Command1_Click 'Exit Sub
  End If
  DrawShape '绘制模仿点阵
  Chinese2HData '横向顺序取模 放入ZMHP中
  PrintHData  '输出ZMHP
  'H2V '对ZMHP纵向顺序取模,放入ZMSP中
  'PrintV  '输出ZMSP
  H2VH '吉林宏源电子 JD12864C 液晶用数据
  PrintVH '打印ZMVH
  'Me.Cls
End Sub



Private Sub Form_Load()
  '**************************************
  On Error Resume Next
  ZitiPath = App.Path & "\" & "hzk16" '
  Text2.Text = "" '横向原数据"
  Text3.Text = "" '横向计算结果"
  Command1.Caption = "生成字模"
  E = 1
  'Text3.Text = "纵向计算结果" & vbCrLf & vbCrLf & "点阵上用的是这个数据"
'  Text5.Text = "纵向取反计算结果"

  If Len(Command) <> 0 Then
    Text1 = Command
    Call Command1_Click
  End If
  
 ' Command1_Click
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


Function DrawShape()

  '=================================
  'Shape1 组成的16*16点阵
  On Error Resume Next
  Dim I2, I, J, TK As Integer
  Dim SI
  
  If N >= 2 Then TK = 2 Else TK = N
  
  For I2 = 0 To TK
    For I = 0 To 15 '16列
      For J = 0 To 15 '16个
        SI = I2 * 256 + I * 16 + J '每行16个
        Load Shape1(SI)

        If I = 8 Then
          Shape1(SI).Top = Shape1((SI Mod 256) - J - 1).Top + 200
        Else
          Shape1(SI).Top = Shape1((SI Mod 256) - J - 1).Top + 120
        End If
        
        If SI Mod 256 = 0 And SI <> 0 Then
          Shape1(SI).Top = Shape1(0).Top
        End If
        
        
        If J = 0 Then
          If SI Mod 256 = 0 Then
            Shape1(SI).Left = Shape1(SI - 1).Left + 500
          Else
            Shape1(SI).Left = Shape1(I2 * 256).Left
          End If
        Else
          If J = 8 Then
            Shape1(SI).Left = Shape1(SI - 1).Left + 200
          Else
            Shape1(SI).Left = Shape1(SI - 1).Left + 120
          End If
        End If
        
        Shape1(SI).Visible = True
      Next
    Next
  Next
  HScroll1.Min = 1
  HScroll1.Value = 1
  HScroll1.Max = N - 2
  If N >= 4 Then HScroll1.Visible = True Else HScroll1.Visible = False
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
  Dim I2, I, J As Integer 'for计数器
  Dim tmpStr As String  '临时字符 存放zmhp(j)
    
  
  For I2 = 1 To N
  
    S = S & vbCrLf & "DB    "
    W = W & vbCrLf & "DB    "

    For J = 1 To 32
      tmpStr = Hex(ZMHP(I2, J))
      If Len(tmpStr) = 1 Then tmpStr = "0" & tmpStr '格式化该字符,长度不为2用0补.
      '==================左侧字符表显示该字符
          
      '显示shape1
      For I = 1 To 8
      If I2 > 3 Then Exit For
       ' 延时Timer1
'        Timer1.Enabled = True
'        While Timer1.Enabled = True
'        DoEvents
'        Wend
        '根据相应位控制相对应Shape1
        
        If (Val("&h" & tmpStr) And 2 ^ (8 - I)) = 0 Then
          Shape1((J - 1) * 8 + I - 1 + (I2 - 1) * 256).BackColor = vbRed
        Else
          Shape1((J - 1) * 8 + I - 1 + (I2 - 1) * 256).BackColor = vbYellow
        End If
      Next
  
      
      S = S & "0" & tmpStr & "H,"
  
      tmpStr = Hex(&HFF - ZMHP(I2, J))
      If Len(tmpStr) = 1 Then tmpStr = "0" & tmpStr
      W = W & "0" & tmpStr & "H,"
  
      If J = 16 Then
        S = S & vbCrLf & "DB    "
        W = W & vbCrLf & "DB    "
      End If
    Next
        S = S & ";" & HzStr1(I2) & vbCrLf
    W = W & vbCrLf ' & "DB "
  Next
  
  Text2 = "横向左右 上下取原数据" & vbCrLf & "(吉林宏源电子  JD240128-7 液晶显示器专用)：" & vbCrLf & S
  'Text3 = "横向取反后数据：" ' & vbCrLf & W

End Function


'
'
''============================================
''将ZMHP 横排字符数组生成数组ZMSP顺序纵向/竖排数组ZMSP
''将ZMHP的1,3,5,7~~~~ 15 的第8位作为ZMSP的第1个
''将ZMHP的17,16,21,23~~~~ 32 的第8位作为ZMSP的第2个
''
''返回结果存放到ZMSP
''
'
'Function H2V()  '将横排转化为竖排
'  Dim I, K, M, M2 As Integer
'  Dim Z, ZA As Byte '临时变量
'  Dim BitI As Byte  '横向取模转纵向取模的第I位, I= 8~1
'  ReDim ZMSP(1 To N, 1 To 32)
'  Dim tmpM22 As Byte '2^
'
'  For I = 1 To N
'    For K = 1 To 2 '
'      For M = 1 To 8
'        Z = 0
'        ZA = 0
'        BitI = 2 ^ (8 - M)
'        For M2 = 0 To 7
'          tmpM22 = 2 ^ (7 - M2)
'          If ((ZMHP(I, K + M2 * 2) And BitI) <> 0) Then Z = Z + tmpM22     '作位最高位
'          If ((ZMHP(I, K + M2 * 2 + 16) And BitI) <> 0) Then ZA = ZA + tmpM22
'        Next
'        ZMSP(I, ((K - 1) * 16) + M * 2 - 1) = Z '取的为上部分
'        ZMSP(I, ((K - 1) * 16) + M * 2) = ZA  '下一部分
'      Next
'    Next
'  Next
'End Function
'
'Function PrintV()
'  Dim I2, I As Integer
'  Dim S, W As String
'  Dim tmpStr As String
'  For I2 = 1 To N
'    S = S & vbCrLf & "DB    "
'    W = W & vbCrLf & "DB    "
'    For I = 1 To 32
'
'      'If I = 1 Then
'      '  S = S & vbCrLf & "DB "
'      '  W = W & vbCrLf & "DB "
'
'      'End If
'      tmpStr = Hex(ZMSP(I2, I))
'      If Len(tmpStr) = 1 Then tmpStr = "0" & tmpStr '格式化,长度变为2 不足前面用0补
'      S = S & "0" & tmpStr & "H,"
'
'      tmpStr = Hex(&HFF - ZMSP(I2, I))
'      If Len(tmpStr) = 1 Then tmpStr = "0" & tmpStr '格式化,长度变为2 不足前面用0补
'      W = W & "0" & tmpStr & "H,"  'Debug.Print Hex(&HFF - zmhp(1, i)) & ",";
'      If I = 16 Then
'        S = S & vbCrLf & "DB "
'        W = W & vbCrLf & "DB "
'      End If
'    Next
'        S = S & ";" & HzStr1(I2) & vbCrLf
'        W = W & vbCrLf
'
'  Next
'
'  Text3.Text = "纵向上下 左右取原数据：" & vbCrLf & S
'  'Text5.Text = "纵向取反后数据：" '& vbCrLf & W
'End Function
'
'
 '============================================
'将ZMHP 横排字符数组生成数组ZMSP顺序纵向/竖排数组ZMSP
'将ZMHP的1,3,5,7~~~~ 15 的第8位作为ZMSP的第1个
'将ZMHP的1,3,5,7~~~~ 15 的第7位作为ZMSP的第2个
'将ZMHP的1,3,5,7~~~~ 15 的第6位作为ZMSP的第3个

'
'返回结果存放到ZMSP
'

Function H2VH()  '将横排转化为竖排
  Dim I, K, M, M2 As Integer
  Dim Z, ZA As Byte '临时变量
  Dim BitI As Byte  '横向取模转纵向取模的第I位, I= 8~1
  ReDim ZMVH(1 To N, 1 To 32)
  Dim tmpM22 As Byte '2^
  
  For I = 1 To N
    For K = 1 To 2 '
      For M = 1 To 8
        Z = 0
        ZA = 0
        BitI = 2 ^ (8 - M)
        For M2 = 0 To 7
          tmpM22 = 2 ^ M2
          If ((ZMHP(I, K + M2 * 2) And BitI) <> 0) Then Z = Z + tmpM22     '作位最高位
          If ((ZMHP(I, K + M2 * 2 + 16) And BitI) <> 0) Then ZA = ZA + tmpM22
        Next
        ZMVH(I, ((K - 1) * 8) + M) = Z  '取的为上部分
        ZMVH(I, ((K - 1) * 8) + M + 16) = ZA '下一部分
      Next
    Next
  Next
End Function

Function PrintVH()
  Dim I2, I As Integer
  Dim S, W As String
  Dim tmpStr As String
  For I2 = 1 To N
    S = S & vbCrLf & "DB    "
    W = W & vbCrLf & "DB    "
    For I = 1 To 32
      
      
      
      'If I = 1 Then
      '  S = S & vbCrLf & "DB "
      '  W = W & vbCrLf & "DB "

      'End If
      tmpStr = Hex(ZMVH(I2, I))
      If Len(tmpStr) = 1 Then tmpStr = "0" & tmpStr '格式化,长度变为2 不足前面用0补
      S = S & "0" & tmpStr & "H,"
      
      tmpStr = Hex(&HFF - ZMVH(I2, I))
      If Len(tmpStr) = 1 Then tmpStr = "0" & tmpStr '格式化,长度变为2 不足前面用0补
      W = W & "0" & tmpStr & "H,"  'Debug.Print Hex(&HFF - zmhp(1, i)) & ",";
      If I = 16 Then
        S = S & vbCrLf & "DB    "
        W = W & vbCrLf & "DB    "
      End If
    Next
        S = S & ";" & HzStr1(I2) & vbCrLf
        W = W & vbCrLf
    
  Next

  Text3.Text = "低位在前,向下左右,上下左右" & vbCrLf & "(吉林宏源电子 JD12864C-1 液晶显示器专用)：" & vbCrLf & S
  'Text5.Text = "纵向取反后数据：" '& vbCrLf & W
End Function

Private Sub HScroll1_Scroll()
  Dim I2, J, I As Integer 'for计数器
  Dim tmpStr As String  '临时字符 存放zmhp(j)
  Timer1.Interval = 0
  Timer1.Enabled = False
  For I2 = 0 To 2

    For J = 1 To 32
      'If I2 > 3 Then Exit For
      tmpStr = Hex(ZMHP(I2 + HScroll1.Value, J))
      
      '显示shape1
      For I = 1 To 8
        If (Val("&h" & tmpStr) And 2 ^ (8 - I)) = 0 Then
          Shape1((J - 1) * 8 + I - 1 + (I2) * 256).BackColor = vbRed
        Else
          Shape1((J - 1) * 8 + I - 1 + (I2) * 256).BackColor = vbYellow
        End If
      Next
    Next
  Next

End Sub

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


