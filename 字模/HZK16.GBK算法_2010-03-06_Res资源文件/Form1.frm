VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "汉字、全角字符点阵提取软件"
   ClientHeight    =   8910
   ClientLeft      =   6960
   ClientTop       =   3600
   ClientWidth     =   9390
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   9390
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame3 
      Caption         =   "字模输出"
      Height          =   2640
      Index           =   3
      Left            =   45
      TabIndex        =   11
      Top             =   6210
      Width           =   9315
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   2235
         Index           =   2
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Text            =   "Form1.frx":030A
         Top             =   315
         Width           =   9105
      End
      Begin VB.CommandButton Command2 
         Caption         =   "复制到剪贴板"
         Height          =   285
         Index           =   2
         Left            =   7605
         TabIndex        =   12
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "字模输出"
      Height          =   2640
      Index           =   2
      Left            =   45
      TabIndex        =   9
      Top             =   3510
      Width           =   9315
      Begin VB.CommandButton Command2 
         Caption         =   "复制到剪贴板"
         Height          =   285
         Index           =   1
         Left            =   7605
         TabIndex        =   4
         Top             =   0
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   2235
         Index           =   1
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Text            =   "Form1.frx":0312
         Top             =   315
         Width           =   9150
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "输入汉字"
      Height          =   750
      Left            =   45
      TabIndex        =   8
      Top             =   45
      Width           =   9315
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   "生成字模"
         Default         =   -1  'True
         Height          =   375
         Left            =   8235
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   225
         Width           =   960
      End
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
         Left            =   90
         TabIndex        =   0
         Top             =   225
         Width           =   8010
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "字体预览"
      Height          =   2625
      Index           =   0
      Left            =   45
      TabIndex        =   7
      Top             =   840
      Width           =   9315
      Begin VB.HScrollBar HScroll1 
         Height          =   240
         Left            =   135
         TabIndex        =   6
         Top             =   2295
         Visible         =   0   'False
         Width           =   8880
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         DrawMode        =   3  'Not Merge Pen
         Height          =   130
         Index           =   0
         Left            =   135
         Shape           =   1  'Square
         Top             =   225
         Visible         =   0   'False
         Width           =   130
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "汉字内码"
      Height          =   1140
      Index           =   1
      Left            =   45
      TabIndex        =   10
      Top             =   4500
      Visible         =   0   'False
      Width           =   9315
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   750
         Index           =   0
         Left            =   45
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   3
         Text            =   "Form1.frx":0318
         Top             =   315
         Width           =   9150
      End
      Begin VB.CommandButton Command2 
         Caption         =   "复制到剪贴板"
         Height          =   285
         Index           =   0
         Left            =   7605
         TabIndex        =   2
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.Menu mnuTray 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuTrayPreView 
         Caption         =   "预　览"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuTrayNeiMa 
         Caption         =   "内　码"
      End
      Begin VB.Menu mnuTray240128 
         Caption         =   "240X128"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnuTray12864 
         Caption         =   "128X64"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuTrayRestore 
         Caption         =   "恢  复"
      End
      Begin VB.Menu MnuTrayLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrayC51 
         Caption         =   "C51 格式"
         Checked         =   -1  'True
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuTrayASM 
         Caption         =   "ASM 格式"
      End
      Begin VB.Menu MnuTrayLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrayAbout 
         Caption         =   "关  于"
      End
      Begin VB.Menu MnuTrayLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrayClose 
         Caption         =   "退  出"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pNid As NotifyIconData) As Long


    Dim HzStr1() As String '从文本中取出的中文存放的一个数组
    Dim N As Integer '文本计数用hzLen()
    Dim ZMHP() As Byte, ZMSP() As Byte, ZMVH() As Byte  '定义横/竖排字存放的树组
    Dim HZK166() As Byte '字库数组
    Dim SWidth As Integer
    Dim Spre As Integer
    Dim SPre8 As Integer
    Dim SPreChr As Integer
    Dim Text1Focus As Boolean
    Public Style As Boolean 'true ASM格式   false c51格式
'======================================================================================================================================


'有效地利用一个API函数 Shell_NotifyIcon和NOTIFYICONDATA数据结构就能达到这一目的，

'
'面的这个程序运行后，将窗口图标加入到了WINDOWS状态栏中，用鼠标右击该图标会弹出一个菜单，
'可实现修改该图标,窗口复位,最小化到系统托盘,最大化及关闭程序等功能
'
'在VB6中新建一工程，将Form1的ScalMode的属性设为3，加入一个image控件和一个对话框控件（要加入对话框控件，
'须在部件中选取Microsoft Common Dialog Control 6.0），将image1的visible属性改为False，为该Form添加一个菜单，菜单设置如下:
'
'标题           名称
'文件(&F)       mnuFile （一级菜单）
'退出(&E)       mnuExit （二级菜单）
'Popup          mnuTray （一级菜单,去掉该项的"可见"项）
'更换图标(&I)   mnuTrayChangeIcon (以下全为二级菜单)
'恢复(&R)       mnuTrayRestore
'最小化(&N)     mnuTrayMinimize"
'最大化(&X)     mnuTrayMaximize
'-              mnuTrayLine
'关闭(&C)       mnuTrayClose

'以下是程序清单:
 '   Private LastState As Integer '保留原窗口状态
    
    '---------- dwMessage可以是以下NIM_ADD、NIM_DELETE、NIM_MODIFY 标识符之一----------
    
    Private Const NIM_ADD = &H0       '在任务栏中增加一个图标
    Private Const NIM_MODIFY = &H1    '修改任务栏中个图标信息
    Private Const NIM_DELETE = &H2    '删除任务栏中的一个图标
    
    Private Const NIF_MESSAGE = &H1   'NOTIFYICONDATA结构中uFlags的控制信息
    Private Const NIF_ICON = &H2
    Private Const NIF_TIP = &H4
    Private Const NIF_INFO = &H10
    
    Private Const WM_MOUSEMOVE = &H200 '当鼠标指针移至图标上
    
    Private Const WM_LBUTTONUP = &H202
    Private Const WM_RBUTTONUP = &H205
    
    Private Type NotifyIconData
        cbSize As Long            '该数据结构的大小
        hwnd As Long              '处理任务栏中图标的窗口句柄
        uID As Long               '定义的任务栏中图标的标识
        uFlags As Long            '任务栏图标功能控制，可以是以下值的组合（一般全包括）
                                  'NIF_MESSAGE 表示发送控制消息；
                                  'NIF_ICON表示显示控制栏中的图标；
                                  'NIF_TIP表示任务栏中的图标有动态提示。
                                  
        UCallbackMessage As Long  '任务栏图标通过它与用户程序交换消息，处理该消息的窗口由hWnd决定
        hIcon As Long             '任务栏中的图标的控制句柄
        szTip As String * 128     '图标的提示信息
        
        
        dwState   As Long
        dwStateMask   As Long
        szInfo   As String * 256
        uTimeoutAndVersion   As Long
        szInfoTitle   As String * 64
        dwInfoFlags   As Long
    End Type
    
    Const niif_info = &H1
    
    Dim myData As NotifyIconData
    Dim Add As Boolean


'====================声明结束

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) '鼠标事件
    X = X / Screen.TwipsPerPixelX
    Select Case CLng(X)
        Case WM_RBUTTONUP '鼠标在图标上右击时弹出菜单
            Me.PopupMenu mnuTray
        Case WM_LBUTTONUP '鼠标在图标上左击时窗口若最小化则恢复窗口位置
            mnuTrayRestore_Click
    End Select
End Sub

Function Adjust() '窗体调整
    Dim i As Integer
    Dim Pos As Integer
    Dim HeightPreFrame As Integer
    HeightPreFrame = 45
    Pos = Frame2.Top + Frame2.Height + HeightPreFrame
    For i = 0 To 3
        If Frame3(i).Visible = True Then
            Frame3(i).Top = Pos
            'MsgBox I & "--->>" & Pos & ">" & Frame3(I).Top & ", " & Frame3(I).Height
            Pos = Pos + Frame3(i).Height + HeightPreFrame
        End If
    Next
    Form1.Height = Pos + 405
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Form_Unload (-1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Shell_NotifyIcon NIM_DELETE, myData '窗口卸载时，将状态栏中的图标一同卸载
End Sub



Private Sub mnuTrayASM_Click() 'ASM格式
    Me.mnuTrayC51.Enabled = True
    Me.mnuTrayC51.Checked = False
    Me.mnuTrayASM.Enabled = False
    Me.mnuTrayASM.Checked = True
    Style = True
    Command1_Click
End Sub

Private Sub mnuTrayC51_Click() 'C51格式
    Me.mnuTrayC51.Enabled = False
    Me.mnuTrayC51.Checked = True
    
    Me.mnuTrayASM.Enabled = True
    Me.mnuTrayASM.Checked = False
    Style = False
    Command1_Click
End Sub

Private Sub mnuTrayNeiMa_Click() '汉字内码控制
    Frame3(1).Visible = Not Frame3(1).Visible
    Me.mnuTrayNeiMa.Checked = Frame3(1).Visible
    Adjust
End Sub

Private Sub mnuTrayPreView_Click() '预览控制
    Frame3(0).Visible = Not Frame3(0).Visible
    Me.mnuTrayPreView.Checked = Frame3(0).Visible
    Adjust
End Sub

Private Sub MnuTray12864_Click() '12864控制
    Frame3(3).Visible = Not Frame3(3).Visible
    Me.MnuTray12864.Checked = Frame3(3).Visible
    Adjust
End Sub

Private Sub mnuTray240128_Click() '240128控制
    Frame3(2).Visible = Not Frame3(2).Visible
    Me.mnuTray240128.Checked = Frame3(2).Visible
    Adjust
End Sub

Private Sub mnuTrayAbout_Click() '关于
    MsgBox Me.Caption & vbCrLf & vbCrLf & "     版本:3.0", vbInformation
End Sub

Private Sub mnuTrayClose_Click() '关闭
    Unload Me
End Sub

Private Sub Form_Resize() '窗体调整大小
    Dim tmpChar As String
    If Me.WindowState = 0 Then
        Me.mnuTrayRestore.Caption = "最小化"
        Me.MnuTray12864.Visible = True
        Me.mnuTray240128.Visible = True
        Me.mnuTrayNeiMa.Visible = True
        Me.mnuTrayPreView.Visible = True
    Else
        Me.MnuTray12864.Visible = False
        Me.mnuTray240128.Visible = False
        Me.mnuTrayNeiMa.Visible = False
        Me.mnuTrayPreView.Visible = False
        Me.Hide
        Me.mnuTrayRestore.Caption = "恢　复"
        TrayShowInfo "软件已经最小化到托盘区!"
    End If
End Sub

Function TrayShowInfo(Info As String)
'显示信息. info为字符串
    With myData
        .cbSize = Len(myData)
        .hwnd = Me.hwnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP Or NIF_INFO  '其中INF_INFO为气泡,如果不需要可以注释掉.
        .UCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon.Handle           '默认为窗口图标
        .szTip = Me.Caption & vbNullChar
        .dwInfoFlags = niif_info
        .szInfoTitle = Me.Caption & "提示您" & vbNullChar
        .szInfo = Info & vbNullChar
    End With


    If Add = True Then
        Shell_NotifyIcon NIM_MODIFY, myData  '修改图标
    Else
        Shell_NotifyIcon NIM_ADD, myData  '加载图标
        Add = True
    End If

End Function


Private Sub mnuTrayRestore_Click()  '恢复菜单
    If Me.WindowState = 0 Then
        Me.WindowState = 1
        Me.Hide
    Else
        Me.WindowState = 0
        Me.Show
        Me.SetFocus
    End If
End Sub

'================================================================================================


Private Sub Frame3_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 0 Then Exit Sub
    Frame3_MouseDown Index, Button, Shift, X, Y
End Sub

Private Sub Frame3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  
    Dim ShapeI As Integer
    Dim HSValue As Integer
    Dim ShapeM As Integer, ShapeN As Integer, ShapeStart As Integer
    Dim i As Integer
    HSValue = HScroll1.Value
    
    Dim SAPE
    For Each SAPE In Shape1
        If X > SAPE.Left And X < SAPE.Width + SAPE.Left And Y > SAPE.Top And Y < SAPE.Top + SAPE.Height Then
'            If SAPE.BackColor = vbRed Then
'                SAPE.BackColor = vbYellow
'            Else
'                SAPE.BackColor = vbRed
'            End If

            
            
            ShapeI = SAPE.Index
            ShapeStart = Int(ShapeI / 8) * 8
            ShapeM = HSValue + Int(ShapeI / 256)
            ShapeN = Int((ShapeI Mod 256) / 8) + 1
'            Debug.Print "----------------------------------"
'            Debug.Print SAPE.Index '; HScroll1.Value; UBound(ZMHP)
'            Debug.Print Int(ShapeI / 8) * 8 '起始位置
'            Debug.Print HSValue + Int(ShapeI / 256); Int((ShapeI Mod 256) / 8) + 1 '下标1和下标2
'            Debug.Print ZMHP(ShapeM, ShapeN)
            ZMHP(ShapeM, ShapeN) = 0
            Select Case Button
            Case 1
                'SAPE.FillColor = vbRed
                SAPE.BackColor = vbRed
            Case 2
                'SAPE.FillColor = vbYellow
                SAPE.BackColor = vbYellow
            End Select
            For i = 0 To 7
                ZMHP(ShapeM, ShapeN) = ZMHP(ShapeM, ShapeN) + (IIf(Shape1(ShapeStart + i).BackColor = vbRed, 0, 1)) * 2 ^ (7 - i)
            Next
            
            ReDraw
'            Debug.Print ZMHP(ShapeM, ShapeN)
            
'            Debug.Print Button
'            Debug.Print SAPE.FillColor
'            Debug.Print SAPE.BackColor

            Exit Sub
        End If
    Next
End Sub

Private Sub Command1_Click()
    SaveChinese
    If N = 0 Then
        Text2(1) = ""
        Text2(2) = ""
        If Len(Text1) <> 0 Then
            Text1 = ""
'            MsgBox "不支持ASCII字符(半角字符)" & vbCrLf & vbCrLf & "支持全角字符及汉字!", vbCritical
        Else
            Text1.SetFocus
        End If
        Exit Sub
    End If

    DrawShape '绘制模仿点阵
    Chinese2HData '横向顺序取模 放入ZMHP中
    ReDraw
End Sub

Function ReDraw() '刷新数据

    PrintHData  '输出ZMHP JD240128-7
    H2VH '吉林宏源电子 JD12864C-1 液晶用数据
    PrintVH '打印ZMVH
    GetNeiMa

End Function
Private Sub Form_Load()
    
    HZK166 = LoadResData(101, "HZK16")
    
    Text2(0).Text = ""
    Text2(1).Text = ""
    Text2(2).Text = ""
    Command1.Caption = "生成字模"
    Frame3(0) = "字体预览"
    Frame3(1) = "汉字内码列表"
    Frame3(2) = "240*128    JD240128-7  吉林市宏源科学仪器有限公司液晶专用"
    Frame3(3) = "128* 64    JD12864C-1  吉林市宏源科学仪器有限公司液晶专用"
    

    If Len(Command) <> 0 Then
        Text1 = Command
        Command1_Click
    End If
    
    TrayShowInfo "软件[" & Me.Caption & "]已运行!"
    
End Sub


Function SaveChinese()
'=======================
'从Text1中找出汉字,存放到HzStr1数组中
'汉字个数存放到 N 中
    ReDim HzStr1(0)
    Dim UboundStr1, i
    For i = 1 To Len(Text1)
        If Asc(Mid(Text1.Text, i, 1)) < 0 Then
            UboundStr1 = UBound(HzStr1) + 1
            ReDim Preserve HzStr1(UboundStr1)
            HzStr1(UboundStr1) = Mid(Text1.Text, i, 1) '把汉字存入数组中.
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
    Dim i As Integer
    Dim J As Integer
    
    
    Spre = 120
    SPre8 = 170
    SPreChr = 200
    
    Shape1(0).Visible = False
    
    For i = 0 To Shape1.UBound
        Unload Shape1(i)
    Next
  
    Dim SWidth As Integer
    Dim Chr As Integer
    
    Chr = IIf(N > 4, 4, N)
    SWidth = Shape1(0).Width

    Dim FrameWidth  As Integer
    Dim ShapeWidth  As Long
    
    FrameWidth = Frame3(0).Width '9285
    ShapeWidth = (((14 * Spre + SWidth + SPre8) * Chr + (Chr - 1) * (SPreChr - SWidth)))
    Shape1(0).Left = (FrameWidth - ShapeWidth) / 2

    TK = IIf(N > 4, 3, N - 1)

    For I2 = 0 To TK
        For i = 0 To 15 '16列
            For J = 0 To 15 '16个
                SI = I2 * 256 + i * 16 + J   '每行16个
                
                Load Shape1(SI)
                
                Shape1(SI).Top = Shape1((SI Mod 256) - J - 1).Top + IIf(i = 8, SPre8, Spre)
                
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
    HScroll1.Max = N - 3
    HScroll1.Visible = IIf(N > 4, True, False)
End Function


'========================
' 汉字顺序横排取模,存放到ZMHP数组中
' 读字模文件存放到Hzk166中
' 找到汉字字模位置,并按行顺序存放到ZMHP(汉字个数,32)数组中
'
Function Chinese2HData() '转化为横排取模
    Dim Address  '存放在hzk16中的地址

    Dim QWM, QM, WM  '文件长度,区位码,区码,位码

    Dim i, J As Integer
    
    ReDim ZMHP(1 To N, 1 To 32) '定义数组 N维
    
    '===================
    '获取汉字区位码
    For i = 1 To UBound(HzStr1)
        QWM = Hex(Asc(HzStr1(i)))  '区位码
        QWM = Right("0000" & QWM, 4)
        QM = Val("&H" & Left(QWM, 2)) '区码
        WM = Val("&H" & Right(QWM, 2)) '位码
        
        Dim X, Y, Z, M

'++++++++++++++++++++++++++++++++++++++++++++++++++
'XYZM的原始表达式
'        If WM > &HA0 Then
'            M = &H5E
'            Y = WM - &HA1
'            If QM > &HA0 Then
'                X = QM - &HA1
'                Z = 0
'            Else
'                X = QM - &H81
'                Z = &H2284
'            End If
'        Else
'            M = &H60
'            If WM > &H7F Then
'                Y = WM - &H41
'            Else
'                Y = WM - &H40
'            End If
'
'            If QM > &HA0 Then
'                X = QM - &HA1
'                Z = &H3A44
'            Else
'                X = QM - &H81
'                Z = &H2E44
'            End If
'        End If
'++++++++++++++++++++++++++++++++++++++++++++++++++
            
        X = QM - IIf(QM > &HA0, &HA1, &H81)
        Y = WM - IIf(WM > &HA0, &HA1, IIf(WM > &H7F, &H41, &H40))
        Z = IIf(WM > &HA0, IIf(QM > &HA0, 0, &H2284), IIf(QM > &HA0, &H3A44, &H2E44))
        M = IIf(WM > &HA0, &H5E, &H60)


        Address = (X * M + Y + Z) * 32
'        Debug.Print Hex(Address)
'--------------------------------------------------------------------
        For J = 1 To 32 '每个字为32个字节
            ZMHP(i, J) = HZK166(Address + J - 1)  '将点阵数据存入,数组
            'Debug.Print ZMHP(i, J)
        Next
    Next
End Function
  
'=============================
'输出横排取模的数组ZMHP数据,并用Shape1 显示
'
'
Function PrintHData() '打印HData
    Dim S As String '存放原码
    Dim I2 As Integer, i As Integer, J As Integer   'for计数器
    Dim tmpStr As String  '临时字符 存放zmhp(j)
    Dim K As Integer
    K = IIf(HScroll1.Value = 1, 1, N - HScroll1.Value + 2) '??为什么+2?
    For I2 = K To N
        If Style Then
            S = S & "DB    "
        Else
            S = S & "{"
        End If
        For J = 1 To 32
            
            tmpStr = Hex(ZMHP(I2, J))
            tmpStr = Right("00" & tmpStr, 2) '格式化该字符,长度不为2用0补.
            '==================左侧字符表显示该字符
            
            '显示shape1
            For i = 1 To 8
                If I2 > 4 Then Exit For
                Shape1((J - 1) * 8 + i - 1 + (I2 - 1) * 256).BackColor = IIf((Val("&h" & tmpStr) And 2 ^ (8 - i)) = 0, GetColor(vbRed), GetColor(vbYellow))
            Next
            If Style Then
                S = S & "0" & tmpStr & "H,"
            Else
                S = S & "0x" & tmpStr & ","
            End If
            If J = 16 Then
                If Style Then
                    S = S & ";" & HzStr1(I2) & vbCrLf & "DB    "
                Else
                    S = S & "},//" & HzStr1(I2) & vbCrLf & "{"
                End If
            End If
        Next
        If Style Then
            S = S & ";" & HzStr1(I2) & vbCrLf
        Else
            S = S & "},//" & HzStr1(I2) & vbCrLf
        End If

    Next
    If Style Then
        Text2(1) = Replace(S, ",;", " ;")
    Else
        Text2(1) = Replace(S, ",},//", "},//")
    End If
End Function

Function H2VH()  '将横排转化为竖排
    Dim i, K, M, M2 As Integer
    Dim Z, ZA As Byte '临时变量
    Dim BitI As Byte  '横向取模转纵向取模的第I位, I= 8~1
    ReDim ZMVH(1 To N, 1 To 32)
    Dim tmpM22 As Byte '=2^
    
    For i = 1 To N
        For K = 1 To 2 '
            For M = 1 To 8
                Z = 0
                ZA = 0
                BitI = 2 ^ (8 - M)
                For M2 = 0 To 7
                    tmpM22 = 2 ^ M2
                    If ((ZMHP(i, K + M2 * 2) And BitI) <> 0) Then Z = Z + tmpM22     '作位最高位
                    If ((ZMHP(i, K + M2 * 2 + 16) And BitI) <> 0) Then ZA = ZA + tmpM22
                Next
                ZMVH(i, ((K - 1) * 8) + M) = Z  '取的为上部分
                ZMVH(i, ((K - 1) * 8) + M + 16) = ZA '下一部分
            Next
        Next
    Next
End Function

Function PrintVH() '打印横排转化为竖排
    Dim I2 As Integer, i As Integer
    Dim S As String, W As String
    Dim tmpStr As String

    For I2 = 1 To N
        If Style Then
            S = S & "DB    "
        Else
            S = S & "{"
        End If
        For i = 1 To 32
            tmpStr = Hex(ZMVH(I2, i))
            tmpStr = Right("00" & tmpStr, 2) '格式化,长度变为2 不足前面用0补
            If Style Then
                S = S & "0" & tmpStr & "H,"
            Else
                S = S & "0x" & tmpStr & ","
            End If
            If i = 16 Then
                If Style Then
                    S = S & ";" & HzStr1(I2) & vbCrLf & "DB    "
                Else
                    S = S & "},//" & HzStr1(I2) & vbCrLf & "{"
                End If
            End If
        Next
        If Style Then
            S = S & ";" & HzStr1(I2) & vbCrLf
        Else
            S = S & "},//" & HzStr1(I2) & vbCrLf
        End If
    Next
    If Style Then
        Text2(2) = Replace(S, ",;", " ;")
    Else
        Text2(2) = Replace(S, ",},//", "},//")
    End If
End Function

Private Sub HScroll1_Change() '预览框架内的水平滚动条
    HScroll1_Scroll
End Sub

Private Sub HScroll1_GotFocus() '预览框架内的水平滚动条
    Text1.SetFocus
End Sub

Private Sub HScroll1_Scroll()
      On Error Resume Next
      Dim I2, J, i As Integer 'for变量
      Dim tmpStr As Byte ' As String  '临时字符 存放zmhp(j)
      For I2 = 0 To 3
            For J = 1 To 32
                tmpStr = ZMHP(I2 + HScroll1.Value, J)
                '显示shape1
                For i = 1 To 8
                    Shape1((J - 1) * 8 + i - 1 + (I2) * 256).BackColor = IIf((tmpStr And 2 ^ (8 - i)) = 0, GetColor(vbRed), GetColor(vbYellow))
                Next
            Next
      Next
      
End Sub


Private Sub Text1_Click()
    'text1 无焦点时点击之后全选
    If Text1Focus = False Then
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1)
    End If
    Text1Focus = True '设置焦点
End Sub

Private Sub command2_Click(Index As Integer)
    
    Clipboard.Clear   ' 清除剪贴板。
    Clipboard.SetText Text2(Index).Text  ' 将正文放置在剪贴板上。
    
End Sub

Function GetColor(Color) '颜色控制
    GetColor = Color
End Function

Function GetNeiMa() '获取汉字内码
    Dim i As Integer
    Dim Line1 As String
    Dim line2 As String
    Dim TmpHex As String
    Line1 = "DB    "
    line2 = ";     "
    For i = 1 To UBound(HzStr1)
        TmpHex = Right("0000" & Hex(Asc(HzStr1(i))), 4)
        Line1 = Line1 & "0" & Left(TmpHex, 2) & "H,"
        Line1 = Line1 & "0" & Right(TmpHex, 2) & "H" & IIf(i = UBound(HzStr1), " ;汉字内码表(高位在前,低位在后)", ",")
        line2 = line2 & HzStr1(i) & IIf(i = UBound(HzStr1), "        ;汉字字符", "        ")
    Next
    Text2(0) = Line1 & vbCrLf & line2
End Function


Private Sub Text1_LostFocus() 'text1 失去焦点
    Text1Focus = False
End Sub



'REWRITE 2009-05-16
'WRITE BY LOVEYU
'WRITE FOR MYSELF

'update 2009-05-27
'Loveyu
'1. Add ClipBoard Button


'update 2009-07-15
'Loveyu
'1. Clear shapes when new chinese input
'2. Increase chinese display in shape from 3 to 7
'3. Autoadjust(From AMEI)  display Position
'4. Debug and optimize

'Update 2009-07-16!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Loveyu
'1. Support HZK16.GBK!!!!!
'2. Optimize!

'Update 2009-07-21
'loveYu
'1. 增加汉字内码
'2. 单个对象变为对象数组

'Update 2009-08-08
'Loveyu
'1.增加托盘图标
'2.动态刷新各Frame位置.
'3.初始界面只包含基本信息

'Updatae 2010 - 3 - 6!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Loveyu
'1.将字库整合到资源文件中,不需要从外部读取字库文件!!!!!!
'2.根据点阵图形重新生成编码!!!!!!
'3.没时间优化程序.存在BUG
'4.版本升到3.0
