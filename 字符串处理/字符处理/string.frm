VERSION 5.00
Begin VB.Form form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Һ���ַ�����"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13035
   Icon            =   "string.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   13035
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   13260
      Top             =   3180
   End
   Begin VB.Frame Cont 
      Caption         =   "��ʾ"
      Height          =   915
      Left            =   4680
      TabIndex        =   11
      Top             =   5640
      Visible         =   0   'False
      Width           =   3195
      Begin VB.Label Lab_str 
         Caption         =   "�ö������Ѹ��Ƶ�������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   360
         Width           =   2595
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   13260
      Top             =   2580
   End
   Begin VB.Frame Frame4 
      Caption         =   "�����DB��"
      Height          =   3735
      Left            =   60
      TabIndex        =   7
      Top             =   4200
      Width           =   12915
      Begin VB.TextBox Text_DB 
         Height          =   3495
         Left            =   60
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   180
         Width           =   12795
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "����"
      Height          =   3135
      Left            =   11460
      TabIndex        =   2
      Top             =   60
      Width           =   1515
      Begin VB.CommandButton command1 
         Caption         =   "ȡ��ģ"
         Height          =   435
         Left            =   120
         TabIndex        =   10
         Top             =   1020
         Width           =   1215
      End
      Begin VB.CommandButton about 
         Caption         =   "����"
         Height          =   435
         Left            =   120
         TabIndex        =   9
         Top             =   1740
         Width           =   1215
      End
      Begin VB.CommandButton Excuect 
         Caption         =   "����"
         Height          =   435
         Left            =   120
         TabIndex        =   4
         Top             =   300
         Width           =   1215
      End
      Begin VB.CommandButton CMD_Exit 
         Caption         =   "�ر�"
         Height          =   435
         Left            =   120
         TabIndex        =   3
         Top             =   2460
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "�ַ�����"
      Height          =   3135
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   11355
      Begin VB.TextBox Text_res 
         Height          =   2895
         Left            =   60
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   11235
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "������ַ�"
      Height          =   915
      Left            =   60
      TabIndex        =   5
      Top             =   3240
      Width           =   12915
      Begin VB.TextBox Text_str 
         Height          =   675
         Left            =   60
         TabIndex        =   6
         Top             =   180
         Width           =   12795
      End
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Timer1.Enabled = False
  Timer2.Enabled = False
  Excuect.Enabled = False
End Sub

'���ڰ�ť��ʾ
Private Sub about_Click()
  MsgBox "03/24/2009 ver 1.3.3 ����""�Զ���DB�����ʼλ��""�������ʾ" & vbCrLf & _
          "03/24/2009 ver 1.3.2 ����""��ť���ı����ڲ�����""" & vbCrLf & _
          "03/24/2009 ver 1.3.1 ����""�Զ���DB����ʼλ������""��ʾ" & vbCrLf & _
          "03/24/2009 ver 1.3   ���""�Զ���DB����ʼλ������""" & vbCrLf & _
          "03/24/2009 ver 1.2   ���""���а����,��ʾ""" & vbCrLf & _
          vbCrLf & _
          "03/23/2009 ver 1.1   ���""ȡ��ģ""" & vbCrLf & _
          "03/23/2009 ver 1.0   ����""����/����""" & vbCrLf & _
          "03/23/2009 ver 0.1   ��ʼ����", vbInformation, "�汾��Ϣ"
End Sub

'�رհ�ť
Private Sub CMD_Exit_Click()
  End
End Sub

'ȡ��ģ��ť
Private Sub Command1_Click()
  On Error Resume Next
  Shell "PCtoLCD2002"
  If Err Then MsgBox "δ�ҵ� PCtoLCD2002 ,��ŵ���ͬĿ¼��!"
End Sub

'����ť
Private Sub Excuect_Click()
'����

  Dim CHNString, DBStr, CHNStr(), STRn, REStr, STR(), STR2() As String
  Dim LenStr, CHN, I, J, LenChn, FixNum As Integer
  On Error Resume Next
  FixNum = Val(InputBox(vbCrLf & "��""ȷ��""��ʹ����������,��""ȡ��""��ʹ��Ĭ��""0""" & vbCrLf & vbCrLf & "(���ַ�Χ: -32767 �� +32767)", "�Զ���DB����ʼλ��", "0"))
  FixNum = FixNum - 1
  If Err Then
    MsgBox "DB�����ʼλ�ó�����Χ" & vbCrLf & vbCrLf & "(���ַ�Χ:" & vbCrLf & "           -32767 �� +32767)", vbOKOnly + 16, "����"
    Exit Sub
  End If
  
  '��ȡ
  REStr = Text_res.Text
  REStr = Trim(REStr)
  LenStr = Len(REStr)
  ReDim STR(LenStr)
  ReDim STR2(LenStr)
  ReDim CHNStr(LenStr)
  CHN = 1
  '���
  For I = 1 To LenStr
    STR(I) = Mid(REStr, I, 1)
    STR2(I) = STR(I)
  Next
  
  'ɾ���Ǻ���
  For I = 1 To LenStr
    If (STR(I) <> Chr(10) And STR(I) <> Chr(13) And Len(STR(I)) = 1) Then
      For J = I + 1 To LenStr
        If STR(I) = STR(J) Then STR(J) = ""
      Next J
    End If
  Next I
  
  '��ȡ����
  For I = 1 To LenStr
    STRn = STR(I)
    If Len(STRn) = 1 And STRn <> Chr(10) And STRn <> Chr(13) Then
      CHNStr(CHN) = STR(I)
      CHN = CHN + 1
    End If
  Next I
  
  '�ϲ� chnstr()����, �����������ַ���
  For I = 1 To LenStr
    CHNString = CHNString + CHNStr(I)
  Next
  CHNString = Trim(CHNString)
  
  '����str2()����  ������ֺ�ĳ�ʼ����
  LenChn = Len(CHNString)
  For I = 1 To LenChn
    If CHNStr(I) <> Chr(10) And CHNStr(I) <> Chr(13) Then
      For J = 1 To LenStr
        If CHNStr(I) = STR2(J) Then STR2(J) = "," & I + FixNum
      Next J
    End If
  Next I
  
  '�ϲ� str2()����
  For I = 1 To LenStr
    DBStr = DBStr + STR2(I)
  Next
  
  '����ϲ����dbstr�ַ���
  DBStr = Trim(DBStr) & ";---"
  
  If Left(DBStr, 1) = "," Then DBStr = "DB " & Right(DBStr, Len(DBStr) - 1)
  
  DBStr = Replace(DBStr, Chr(13) & Chr(10), ";---" & Chr(13) & Chr(10))
  DBStr = Replace(DBStr, Chr(13) & Chr(10) & ",", Chr(13) & Chr(10) & "DB ")
  
  '���� restr�ַ��� ԭʼ�ַ���
  REStr = Replace(REStr, Chr(13) & Chr(10), "---" & Chr(13) & Chr(10))
  
  'split �ַ��� dbstr, restr
  SP1 = Split(DBStr, "---")
  SP2 = Split(REStr, "---")

  DBStr = ""
  
  For I = 0 To UBound(SP1)
    DBStr = DBStr & Trim(SP1(I)) & Chr(9) & Trim(Replace(SP2(I), Chr(13) & Chr(10), ""))
  Next
  
  '��ʾ
  Text_str.Text = Space(FixNum + 1) & CHNString
  Text_DB.Text = DBStr
End Sub

'�ַ������ı���
Private Sub Text_res_Change()
  If Len(Text_res.Text) > 0 Then
    Excuect.Enabled = True
  Else
    Excuect.Enabled = False
  End If
End Sub

'�ַ������ı���
Private Sub Text_str_Click()
'Clipboard.Clear
'Clipboard.SetText Text_str.Text
'Cont.Top = 3240
'Timer1.Enabled = True
'Timer2.Enabled = True
End Sub

'�����db���ı���
Private Sub Text_DB_Click()
'Clipboard.Clear
'Clipboard.SetText Text_DB.Text
'Cont.Top = 5640
'Timer1.Enabled = True
'Timer2.Enabled = True
End Sub

'��ʾ��˸�����ʱ��
Private Sub Timer1_Timer()
  If Cont.Visible = True Then
    Cont.Visible = False
  Else
    Cont.Visible = True
  End If
End Sub

'��ʾ��˸ʱ���ܳ���ʱ��
Private Sub Timer2_Timer()
  Timer1.Enabled = False
  Timer2.Enabled = False
  Cont.Visible = False
End Sub
