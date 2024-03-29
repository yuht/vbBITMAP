VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "汉字字模"
   ClientHeight    =   8505
   ClientLeft      =   6960
   ClientTop       =   3600
   ClientWidth     =   9405
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8505
   ScaleMode       =   0  'User
   ScaleWidth      =   9405
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame3 
      Caption         =   "字模输出"
      Height          =   5610
      Left            =   45
      TabIndex        =   9
      Top             =   2835
      Width           =   9315
      Begin VB.CommandButton Command3 
         Caption         =   "Command3 复制到剪贴板"
         Height          =   375
         Left            =   7575
         TabIndex        =   5
         Top             =   2880
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2 复制到剪贴板"
         Height          =   375
         Left            =   7620
         TabIndex        =   3
         Top             =   180
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   2190
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Text            =   "Form1.frx":030A
         Top             =   630
         Width           =   9150
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   2205
         Left            =   45
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Text            =   "Form1.frx":0310
         Top             =   3330
         Width           =   9150
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   435
         Left            =   75
         TabIndex        =   11
         Top             =   2880
         Width           =   7275
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   435
         Left            =   90
         TabIndex        =   10
         Top             =   180
         Width           =   7275
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "输入汉字"
      Height          =   750
      Left            =   45
      TabIndex        =   8
      Top             =   45
      Width           =   9315
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   0
         Top             =   240
         Width           =   7455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1 生成字模"
         Default         =   -1  'True
         Height          =   375
         Left            =   7740
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "字体预览"
      Height          =   1905
      Left            =   45
      TabIndex        =   7
      Top             =   855
      Width           =   9285
      Begin VB.HScrollBar HScroll1 
         Height          =   240
         Left            =   180
         TabIndex        =   2
         Top             =   1530
         Visible         =   0   'False
         Width           =   8880
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         DrawMode        =   3  'Not Merge Pen
         Height          =   90
         Index           =   0
         Left            =   135
         Shape           =   1  'Square
         Top             =   225
         Visible         =   0   'False
         Width           =   90
      End
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

  Dim SWidth As Integer
  Dim Spre As Integer
  Dim SPre8 As Integer
  Dim SPreChr As Integer


Private Sub Command1_Click()
  Me.Cls
  SaveChinese '从 text1中找出汉字
  
  If N = 0 Then
    Text1 = "于海涛"
    Text1.SetFocus
    Command1_Click 'Exit Sub
  End If
  DrawShape '绘制模仿点阵
  Chinese2HData '横向顺序取模 放入ZMHP中
  PrintHData  '输出ZMHP JD240128-7
  H2VH '吉林宏源电子 JD12864C-1 液晶用数据
  PrintVH '打印ZMVH

End Sub





Private Sub Form_Load()
  
  On Error Resume Next
  ZitiPath = App.Path & "\" & "hzk16" '
  Text2.Text = ""
  Text3.Text = ""
  Command1.Caption = "生成字模"
  Command2.Caption = "复制到剪贴板"
  Command3.Caption = "复制到剪贴板"
'  Label1 = "横向左右 上下取原数据" & vbCrLf & ""
 Label1 = vbCrLf & "(吉林宏源电子  JD240128-7 液晶显示器专用):"
'  Label2 = "低位在前,向下左右,上下左右" & vbCrLf & "(吉林宏源电子 JD12864C-1 液晶显示器专用):"
 Label2 = vbCrLf & "(吉林宏源电子  JD12864C-1 液晶显示器专用):"

  If Len(Command) <> 0 Then
    Text1 = Command
    Call Command1_Click
  End If
  
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
    
    
    
    Dim SI As Integer
    Dim TK As Integer
    Dim I2 As Integer
    Dim I As Integer
    Dim J As Integer
    
    
    Spre = 75
    SPre8 = 90
    SPreChr = 150
  
    For I = 0 To Shape1.UBound
        Unload Shape1(I)
    Next
  
    Dim SWidth As Integer
      
    SWidth = Shape1(0).Width
    Dim Chr As Integer
    Chr = IIf(N > 7, 7, N)
    Shape1(0).Left = (Frame1.Width - ((14 * Spre + SWidth + SPre8) * Chr + (Chr - 1) * (SPreChr - SWidth))) / 2
    
  
    TK = IIf(N > 7, 6, N - 1)
    
    
    
    For I2 = 0 To TK
        For I = 0 To 15 '16列
            For J = 0 To 15 '16个
                SI = I2 * 256 + I * 16 + J   '每行16个
                
                Load Shape1(SI)
                
                Shape1(SI).Top = Shape1((SI Mod 256) - J - 1).Top + IIf(I = 8, SPre8, Spre)
                
                If SI Mod 256 = 0 And SI <> 0 Then Shape1(SI).Top = Shape1(0).Top
                
                If J = 0 Then
                    Shape1(SI).Left = IIf(SI Mod 256 = 0, Shape1(SI - 1).Left + SPreChr, Shape1(I2 * 256).Left)
                Else
                    Shape1(SI).Left = Shape1(SI - 1).Left + IIf(J = 8, SPre8, Spre)
                End If
                
                Shape1(SI).Visible = True
            
            Next
        Next
    Next
    HScroll1.Min = 1
    HScroll1.Value = 1
    HScroll1.Max = N - 6
    HScroll1.Visible = IIf(N > 7, True, False)
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
          
      If Sum = 0 Then
            Close #intFileNum
            Kill FileNa
            MsgBox "请将汉字库文件hzk16(748 KB (766,080 字节) 或 261 KB (267,616 字节)) 放入当前目录中!", vbExclamation
            End
      End If
      
      
      ReDim Hzk166(1 To Sum)  '按文件长度定义数组
      Get #intFileNum, , Hzk166 '存入到数组 Hzk166
      Close #intFileNum '关闭字库文件,防止发生错误
      
      ReDim ZMHP(1 To N, 1 To 32) '定义数组 N维
    
      '===================
      '获取汉字区位码
      For I = 1 To UBound(HzStr1)
            QWM = Hex(Asc(HzStr1(I)) - &HA0A0) '区位码
            QWM = Right("0000" & QWM, 4)
            QM = Left(QWM, 2) '区码
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
    Dim S As String '存放原码
    Dim I2, I, J As Integer 'for计数器
    Dim tmpStr As String  '临时字符 存放zmhp(j)
    
    For I2 = 1 To N
        S = S & "DB    "
        For J = 1 To 32
            tmpStr = Hex(ZMHP(I2, J))
            tmpStr = Right("00" & tmpStr, 2) '格式化该字符,长度不为2用0补.
            '==================左侧字符表显示该字符
            
            '显示shape1
            For I = 1 To 8
                If I2 > 7 Then Exit For
                Shape1((J - 1) * 8 + I - 1 + (I2 - 1) * 256).BackColor = IIf((Val("&h" & tmpStr) And 2 ^ (8 - I)) = 0, GetColor(vbRed), GetColor(vbYellow))
            Next
        
            S = S & "0" & tmpStr & "H,"
            
            If J = 16 Then
                S = S & ";" & HzStr1(I2) & vbCrLf & "DB    "
            End If
        Next
        S = S & ";" & HzStr1(I2) & vbCrLf

    Next
    Text2 = Replace(S, ",;", " ;")
End Function

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
        S = S & "DB    "
        For I = 1 To 32
            tmpStr = Hex(ZMVH(I2, I))
            tmpStr = Right("00" & tmpStr, 2) '格式化,长度变为2 不足前面用0补
            S = S & "0" & tmpStr & "H,"
            If I = 16 Then
                S = S & ";" & HzStr1(I2) & vbCrLf & "DB    "
            End If
        Next
        S = S & ";" & HzStr1(I2) & vbCrLf
    Next
    Text3.Text = Replace(S, ",;", " ;")
End Function

Private Sub HScroll1_Change()
  HScroll1_Scroll
End Sub

Private Sub HScroll1_Scroll()
      On Error Resume Next
      Dim I2, J, I As Integer 'for变量
      Dim tmpStr As String  '临时字符 存放zmhp(j)
      For I2 = 0 To 6
            For J = 1 To 32
                tmpStr = Hex(ZMHP(I2 + HScroll1.Value, J))
                '显示shape1
                For I = 1 To 8
                    Shape1((J - 1) * 8 + I - 1 + (I2) * 256).BackColor = IIf((Val("&h" & tmpStr) And 2 ^ (8 - I)) = 0, GetColor(vbRed), GetColor(vbYellow))
                Next
            Next
      Next
End Sub
 


Private Sub Text1_Click()
    '点击之后全选
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)
End Sub

Private Sub command2_Click()
    
    Clipboard.Clear   ' 清除剪贴板。
    Clipboard.SetText Text2.Text  ' 将正文放置在剪贴板上。
End Sub

Private Sub command3_Click()
    Clipboard.Clear   ' 清除剪贴板。
    Clipboard.SetText Text3.Text  ' 将正文放置在剪贴板上。
End Sub

Function GetColor(Color)
    GetColor = Color
End Function

'REWRITE 2009-05-16
'WRITE BY LOVEYU
'WRITE FOR MYSELF

'update 2009-05-27
'Loveyu
'1. add clipboard button


'update 2009-07-15
'Loveyu
'2. Clear shapes when new chinese input
'3. Increase chinese display in shape from 3 to 7
'4. Autoadjust (From AMEI)  display Position
'5. Debug and Better

