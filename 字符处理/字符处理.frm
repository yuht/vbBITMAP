VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form__Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "字符处理"
   ClientHeight    =   8295
   ClientLeft      =   4710
   ClientTop       =   2700
   ClientWidth     =   10155
   Icon            =   "字符处理.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   10155
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton HotKey_D 
      Caption         =   "&D"
      Height          =   255
      Left            =   10260
      TabIndex        =   27
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton HotKey_F 
      Caption         =   "&F"
      Height          =   255
      Left            =   10260
      TabIndex        =   26
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton HotKey_I 
      Caption         =   "&I"
      Height          =   255
      Left            =   10260
      TabIndex        =   25
      Top             =   1140
      Width           =   855
   End
   Begin VB.Frame Frame_Info 
      Caption         =   "提示"
      Height          =   795
      Left            =   3660
      TabIndex        =   16
      Top             =   3600
      Visible         =   0   'False
      Width           =   3855
      Begin VB.TextBox Text_GF 
         Height          =   210
         Left            =   4140
         TabIndex        =   18
         Top             =   420
         Width           =   105
      End
      Begin VB.Label Label_Info 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   60
         TabIndex        =   17
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.Frame Frame_S 
      Caption         =   " 设 置 "
      Height          =   8175
      Left            =   60
      TabIndex        =   11
      Top             =   60
      Visible         =   0   'False
      Width           =   10035
      Begin MSComDlg.CommonDialog CommonDialog 
         Left            =   8400
         Top             =   1380
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "*.* ,*.txt"
      End
      Begin VB.Frame Frame_UseFL 
         Caption         =   "外部字库"
         Height          =   1515
         Left            =   300
         TabIndex        =   19
         Top             =   1140
         Width           =   7575
         Begin VB.CommandButton Command_FL_Bro 
            Appearance      =   0  'Flat
            Caption         =   "...(&B)"
            Height          =   255
            Left            =   6660
            TabIndex        =   28
            Top             =   1020
            Width           =   675
         End
         Begin VB.TextBox Text_FL_Path 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   375
            Left            =   180
            TabIndex        =   21
            Top             =   960
            Width           =   7215
         End
         Begin VB.CheckBox Check_UseFL 
            Caption         =   " 是否生成字模"
            Height          =   315
            Left            =   180
            TabIndex        =   20
            Top             =   240
            Width           =   1635
         End
         Begin VB.Label Label_Help_FL 
            Alignment       =   2  'Center
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7020
            TabIndex        =   24
            Top             =   660
            Width           =   495
         End
         Begin VB.Label Label_FL_Path_Str 
            Caption         =   "输入字模文件所在路径:"
            Enabled         =   0   'False
            Height          =   255
            Left            =   180
            TabIndex        =   23
            Top             =   600
            Width           =   2055
         End
      End
      Begin VB.Frame Frame_DB_Num 
         Caption         =   "DB表起始号码"
         Height          =   615
         Left            =   300
         TabIndex        =   13
         Top             =   360
         Width           =   7575
         Begin VB.TextBox Text_DB_Num 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4020
            MaxLength       =   11
            TabIndex        =   14
            Text            =   "0"
            Top             =   180
            Width           =   2775
         End
         Begin VB.Label Label_DB_Prompt 
            Alignment       =   2  'Center
            Caption         =   "输入自定义的DB表起始数字(默认为0):"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   240
            Width           =   3735
         End
         Begin VB.Label Label_Help_DB_Num 
            Alignment       =   2  'Center
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7020
            TabIndex        =   15
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.CommandButton Command_S_ok 
         Caption         =   "确 定(&O)"
         Height          =   375
         Left            =   8700
         TabIndex        =   12
         Top             =   180
         Width           =   1215
      End
   End
   Begin VB.Frame Frame_Control 
      Caption         =   "控制"
      Height          =   3555
      Left            =   8640
      TabIndex        =   6
      Top             =   60
      Width           =   1455
      Begin VB.CommandButton Command_S 
         Caption         =   "设 置(&S)"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton Command_End 
         Caption         =   "关 闭(&C)"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   3060
         Width           =   1215
      End
   End
   Begin VB.Frame Frame_Output_DB 
      Caption         =   "DB码表(D)"
      Height          =   3555
      Left            =   60
      TabIndex        =   2
      Top             =   4680
      Width           =   10035
      Begin VB.TextBox Text_DB 
         Appearance      =   0  'Flat
         Height          =   3315
         Left            =   60
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   180
         Width           =   9555
      End
      Begin VB.Label Label_CDB 
         Alignment       =   2  'Center
         Caption         =   "               复制"
         Height          =   3195
         Left            =   9720
         TabIndex        =   10
         Top             =   240
         Width           =   195
      End
   End
   Begin VB.Frame Frame_Output_Str 
      Caption         =   "字符串(F)"
      Height          =   975
      Left            =   60
      TabIndex        =   1
      Top             =   3660
      Width           =   10035
      Begin VB.TextBox Text_Str 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   60
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   180
         Width           =   9555
      End
      Begin VB.Label Label_CStr 
         Alignment       =   2  'Center
         Caption         =   " 复制"
         Height          =   735
         Left            =   9720
         TabIndex        =   9
         Top             =   180
         Width           =   195
      End
   End
   Begin VB.Frame Frame_Input_Res 
      Caption         =   "输入字符(I)"
      Height          =   3555
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   8535
      Begin VB.TextBox Text_Res 
         Appearance      =   0  'Flat
         Height          =   3315
         Left            =   60
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   180
         Width           =   8415
      End
   End
End
Attribute VB_Name = "Form__Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check_UseFL_Click()
  If Check_UseFL.Value Then
    Text_FL_Path.Enabled = True
    Text_FL_Path.BackColor = &H80000005
    Label_FL_Path_Str.Enabled = True
  Else
    Text_FL_Path.Enabled = False
    Text_FL_Path.BackColor = &H8000000F
    Label_FL_Path_Str.Enabled = False
  End If
End Sub

Private Sub Command_end_Click()
  End
End Sub

Private Sub Command_FL_Bro_Click()
  On Error Resume Next
    With CommonDialog
    .CancelError = True                 ' 设置标志 When this property is set to True, error number 32755 (cdCancel) occurs whenever the user chooses the Cancel button.
    .InitDir = App.Path                 ' 默认的文件夹为当前文件夹
    .Flags = cdlOFNHideReadOnly         ' 设置过滤器
    .Filter = "文件图标(*.ico)|*.ico" '    "图标文件 (*.ico),*.ico"   ' 指定缺省的过滤器为图标文件
    .ShowOpen                           ' 显示选定文件的名字
  End With
  

  Text_FL_Path = CommonDialog.FileName
End Sub

Private Sub Command_S_ok_Click()
  Frame_S.Visible = False
End Sub

Private Sub Command_S_Click()
  Frame_S.Visible = True
End Sub

Private Sub HOtkey_F_Click()
  Text_Str.SetFocus
End Sub

Private Sub HOTkey_I_Click()
  Text_Res.SetFocus
End Sub

Private Sub HotKey_D_Click()
  Text_DB.SetFocus
End Sub

Private Sub Label_CDB_Click() 'DB码表复制提示
  Clipboard.Clear
  Clipboard.SetText Text_DB.Text
  InfoBox ("DB表已复制到剪切板.")
End Sub

Private Sub Label_CStr_Click() '字符串复制提示
  Clipboard.Clear
  Clipboard.SetText Text_Str.Text
  InfoBox ("字符串已复制到剪切板.")
End Sub

Private Sub Label_Help_DB_Num_Click() '自定义DB码表起始位置提示
  InfoBox "数字介于 -2147483648 和 2147483647."
End Sub

Private Sub Label_Help_FL_Click() '调用字模提示
  InfoBox ("使用外部字库将在DB码表后生成码表对应字模")
End Sub

Private Sub Text_DB_Num_Change() '用文本实现的一个long型的数字判断函数. 累
  
  Dim Num As Long
  Dim Str As String
  
  Str = Trim(Text_DB_Num.Text)
  
  If Len(Str) = 0 Or Str = "-" Then Exit Sub
  
  
  If Left(Str, 1) = "-" Then
    If Len(Str) = 11 Then
      If Right(Str, Len(Str) - 1) > "2147483648" Then
        Text_DB_Num.Text = "-2147483648"
      End If
    End If
  Else
    If Len(Str) > 9 Then
      If Str > "2147483647" Then
        Text_DB_Num.Text = "2147483647"
      End If
    End If
  End If
  
  Num = CLng(Text_DB_Num.Text)
  
  
End Sub


Private Sub Text_FL_Path_LostFocus()  '检测文件是否存在
  On Error Resume Next
  Open Text_FL_Path.Text For Input Access Read As #1  '打开文件,文件号为1#
  Close #1  '关闭1#文件
  If Err Then
    InfoBox (Err.Description)
    Err.Clear
  End If
End Sub

Private Sub Text_GF_LostFocus() '提示frame中的text_gf 文本框(getfocus), 失去焦点时
  Frame_Info.Visible = False
End Sub

Private Sub Label_Info_Click() '点击Label_Info后Frame_Info不可见
  Frame_Info.Visible = False
End Sub

Private Sub Frame_Info_Click() '点击Frame_Info后Frame_Info不可见
  Frame_Info.Visible = False
End Sub

Function InfoBox(Info As String)  '提示信息为info
    Frame_Info.Visible = True
    Text_GF.SetFocus
    Label_Info.Caption = Info
End Function

