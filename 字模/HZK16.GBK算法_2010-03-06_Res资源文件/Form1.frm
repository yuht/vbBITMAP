VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���֡�ȫ���ַ�������ȡ���"
   ClientHeight    =   8910
   ClientLeft      =   6960
   ClientTop       =   3600
   ClientWidth     =   9390
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   9390
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame Frame3 
      Caption         =   "��ģ���"
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
         Caption         =   "���Ƶ�������"
         Height          =   285
         Index           =   2
         Left            =   7605
         TabIndex        =   12
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "��ģ���"
      Height          =   2640
      Index           =   2
      Left            =   45
      TabIndex        =   9
      Top             =   3510
      Width           =   9315
      Begin VB.CommandButton Command2 
         Caption         =   "���Ƶ�������"
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
      Caption         =   "���뺺��"
      Height          =   750
      Left            =   45
      TabIndex        =   8
      Top             =   45
      Width           =   9315
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   "������ģ"
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
            Name            =   "����"
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
      Caption         =   "����Ԥ��"
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
      Caption         =   "��������"
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
         Caption         =   "���Ƶ�������"
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
         Caption         =   "Ԥ����"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuTrayNeiMa 
         Caption         =   "�ڡ���"
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
         Caption         =   "��  ��"
      End
      Begin VB.Menu MnuTrayLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrayC51 
         Caption         =   "C51 ��ʽ"
         Checked         =   -1  'True
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuTrayASM 
         Caption         =   "ASM ��ʽ"
      End
      Begin VB.Menu MnuTrayLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrayAbout 
         Caption         =   "��  ��"
      End
      Begin VB.Menu MnuTrayLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrayClose 
         Caption         =   "��  ��"
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


    Dim HzStr1() As String '���ı���ȡ�������Ĵ�ŵ�һ������
    Dim N As Integer '�ı�������hzLen()
    Dim ZMHP() As Byte, ZMSP() As Byte, ZMVH() As Byte  '�����/�����ִ�ŵ�����
    Dim HZK166() As Byte '�ֿ�����
    Dim SWidth As Integer
    Dim Spre As Integer
    Dim SPre8 As Integer
    Dim SPreChr As Integer
    Dim Text1Focus As Boolean
    Public Style As Boolean 'true ASM��ʽ   false c51��ʽ
'======================================================================================================================================


'��Ч������һ��API���� Shell_NotifyIcon��NOTIFYICONDATA���ݽṹ���ܴﵽ��һĿ�ģ�

'
'�������������к󣬽�����ͼ����뵽��WINDOWS״̬���У�������һ���ͼ��ᵯ��һ���˵���
'��ʵ���޸ĸ�ͼ��,���ڸ�λ,��С����ϵͳ����,��󻯼��رճ���ȹ���
'
'��VB6���½�һ���̣���Form1��ScalMode��������Ϊ3������һ��image�ؼ���һ���Ի���ؼ���Ҫ����Ի���ؼ���
'���ڲ�����ѡȡMicrosoft Common Dialog Control 6.0������image1��visible���Ը�ΪFalse��Ϊ��Form���һ���˵����˵���������:
'
'����           ����
'�ļ�(&F)       mnuFile ��һ���˵���
'�˳�(&E)       mnuExit �������˵���
'Popup          mnuTray ��һ���˵�,ȥ�������"�ɼ�"�
'����ͼ��(&I)   mnuTrayChangeIcon (����ȫΪ�����˵�)
'�ָ�(&R)       mnuTrayRestore
'��С��(&N)     mnuTrayMinimize"
'���(&X)     mnuTrayMaximize
'-              mnuTrayLine
'�ر�(&C)       mnuTrayClose

'�����ǳ����嵥:
 '   Private LastState As Integer '����ԭ����״̬
    
    '---------- dwMessage����������NIM_ADD��NIM_DELETE��NIM_MODIFY ��ʶ��֮һ----------
    
    Private Const NIM_ADD = &H0       '��������������һ��ͼ��
    Private Const NIM_MODIFY = &H1    '�޸��������и�ͼ����Ϣ
    Private Const NIM_DELETE = &H2    'ɾ���������е�һ��ͼ��
    
    Private Const NIF_MESSAGE = &H1   'NOTIFYICONDATA�ṹ��uFlags�Ŀ�����Ϣ
    Private Const NIF_ICON = &H2
    Private Const NIF_TIP = &H4
    Private Const NIF_INFO = &H10
    
    Private Const WM_MOUSEMOVE = &H200 '�����ָ������ͼ����
    
    Private Const WM_LBUTTONUP = &H202
    Private Const WM_RBUTTONUP = &H205
    
    Private Type NotifyIconData
        cbSize As Long            '�����ݽṹ�Ĵ�С
        hwnd As Long              '������������ͼ��Ĵ��ھ��
        uID As Long               '�������������ͼ��ı�ʶ
        uFlags As Long            '������ͼ�깦�ܿ��ƣ�����������ֵ����ϣ�һ��ȫ������
                                  'NIF_MESSAGE ��ʾ���Ϳ�����Ϣ��
                                  'NIF_ICON��ʾ��ʾ�������е�ͼ�ꣻ
                                  'NIF_TIP��ʾ�������е�ͼ���ж�̬��ʾ��
                                  
        UCallbackMessage As Long  '������ͼ��ͨ�������û����򽻻���Ϣ���������Ϣ�Ĵ�����hWnd����
        hIcon As Long             '�������е�ͼ��Ŀ��ƾ��
        szTip As String * 128     'ͼ�����ʾ��Ϣ
        
        
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


'====================��������

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) '����¼�
    X = X / Screen.TwipsPerPixelX
    Select Case CLng(X)
        Case WM_RBUTTONUP '�����ͼ�����һ�ʱ�����˵�
            Me.PopupMenu mnuTray
        Case WM_LBUTTONUP '�����ͼ�������ʱ��������С����ָ�����λ��
            mnuTrayRestore_Click
    End Select
End Sub

Function Adjust() '�������
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
    Shell_NotifyIcon NIM_DELETE, myData '����ж��ʱ����״̬���е�ͼ��һͬж��
End Sub



Private Sub mnuTrayASM_Click() 'ASM��ʽ
    Me.mnuTrayC51.Enabled = True
    Me.mnuTrayC51.Checked = False
    Me.mnuTrayASM.Enabled = False
    Me.mnuTrayASM.Checked = True
    Style = True
    Command1_Click
End Sub

Private Sub mnuTrayC51_Click() 'C51��ʽ
    Me.mnuTrayC51.Enabled = False
    Me.mnuTrayC51.Checked = True
    
    Me.mnuTrayASM.Enabled = True
    Me.mnuTrayASM.Checked = False
    Style = False
    Command1_Click
End Sub

Private Sub mnuTrayNeiMa_Click() '�����������
    Frame3(1).Visible = Not Frame3(1).Visible
    Me.mnuTrayNeiMa.Checked = Frame3(1).Visible
    Adjust
End Sub

Private Sub mnuTrayPreView_Click() 'Ԥ������
    Frame3(0).Visible = Not Frame3(0).Visible
    Me.mnuTrayPreView.Checked = Frame3(0).Visible
    Adjust
End Sub

Private Sub MnuTray12864_Click() '12864����
    Frame3(3).Visible = Not Frame3(3).Visible
    Me.MnuTray12864.Checked = Frame3(3).Visible
    Adjust
End Sub

Private Sub mnuTray240128_Click() '240128����
    Frame3(2).Visible = Not Frame3(2).Visible
    Me.mnuTray240128.Checked = Frame3(2).Visible
    Adjust
End Sub

Private Sub mnuTrayAbout_Click() '����
    MsgBox Me.Caption & vbCrLf & vbCrLf & "     �汾:3.0", vbInformation
End Sub

Private Sub mnuTrayClose_Click() '�ر�
    Unload Me
End Sub

Private Sub Form_Resize() '���������С
    Dim tmpChar As String
    If Me.WindowState = 0 Then
        Me.mnuTrayRestore.Caption = "��С��"
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
        Me.mnuTrayRestore.Caption = "�֡���"
        TrayShowInfo "����Ѿ���С����������!"
    End If
End Sub

Function TrayShowInfo(Info As String)
'��ʾ��Ϣ. infoΪ�ַ���
    With myData
        .cbSize = Len(myData)
        .hwnd = Me.hwnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP Or NIF_INFO  '����INF_INFOΪ����,�������Ҫ����ע�͵�.
        .UCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon.Handle           'Ĭ��Ϊ����ͼ��
        .szTip = Me.Caption & vbNullChar
        .dwInfoFlags = niif_info
        .szInfoTitle = Me.Caption & "��ʾ��" & vbNullChar
        .szInfo = Info & vbNullChar
    End With


    If Add = True Then
        Shell_NotifyIcon NIM_MODIFY, myData  '�޸�ͼ��
    Else
        Shell_NotifyIcon NIM_ADD, myData  '����ͼ��
        Add = True
    End If

End Function


Private Sub mnuTrayRestore_Click()  '�ָ��˵�
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
'            Debug.Print Int(ShapeI / 8) * 8 '��ʼλ��
'            Debug.Print HSValue + Int(ShapeI / 256); Int((ShapeI Mod 256) / 8) + 1 '�±�1���±�2
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
'            MsgBox "��֧��ASCII�ַ�(����ַ�)" & vbCrLf & vbCrLf & "֧��ȫ���ַ�������!", vbCritical
        Else
            Text1.SetFocus
        End If
        Exit Sub
    End If

    DrawShape '����ģ�µ���
    Chinese2HData '����˳��ȡģ ����ZMHP��
    ReDraw
End Sub

Function ReDraw() 'ˢ������

    PrintHData  '���ZMHP JD240128-7
    H2VH '���ֺ�Դ���� JD12864C-1 Һ��������
    PrintVH '��ӡZMVH
    GetNeiMa

End Function
Private Sub Form_Load()
    
    HZK166 = LoadResData(101, "HZK16")
    
    Text2(0).Text = ""
    Text2(1).Text = ""
    Text2(2).Text = ""
    Command1.Caption = "������ģ"
    Frame3(0) = "����Ԥ��"
    Frame3(1) = "���������б�"
    Frame3(2) = "240*128    JD240128-7  �����к�Դ��ѧ�������޹�˾Һ��ר��"
    Frame3(3) = "128* 64    JD12864C-1  �����к�Դ��ѧ�������޹�˾Һ��ר��"
    

    If Len(Command) <> 0 Then
        Text1 = Command
        Command1_Click
    End If
    
    TrayShowInfo "���[" & Me.Caption & "]������!"
    
End Sub


Function SaveChinese()
'=======================
'��Text1���ҳ�����,��ŵ�HzStr1������
'���ָ�����ŵ� N ��
    ReDim HzStr1(0)
    Dim UboundStr1, i
    For i = 1 To Len(Text1)
        If Asc(Mid(Text1.Text, i, 1)) < 0 Then
            UboundStr1 = UBound(HzStr1) + 1
            ReDim Preserve HzStr1(UboundStr1)
            HzStr1(UboundStr1) = Mid(Text1.Text, i, 1) '�Ѻ��ִ���������.
        End If
    Next
    N = UboundStr1
End Function


Function DrawShape()

    '=================================
    'Shape1 ��ɵ�16*16����
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
        For i = 0 To 15 '16��
            For J = 0 To 15 '16��
                SI = I2 * 256 + i * 16 + J   'ÿ��16��
                
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
' ����˳�����ȡģ,��ŵ�ZMHP������
' ����ģ�ļ���ŵ�Hzk166��
' �ҵ�������ģλ��,������˳���ŵ�ZMHP(���ָ���,32)������
'
Function Chinese2HData() 'ת��Ϊ����ȡģ
    Dim Address  '�����hzk16�еĵ�ַ

    Dim QWM, QM, WM  '�ļ�����,��λ��,����,λ��

    Dim i, J As Integer
    
    ReDim ZMHP(1 To N, 1 To 32) '�������� Nά
    
    '===================
    '��ȡ������λ��
    For i = 1 To UBound(HzStr1)
        QWM = Hex(Asc(HzStr1(i)))  '��λ��
        QWM = Right("0000" & QWM, 4)
        QM = Val("&H" & Left(QWM, 2)) '����
        WM = Val("&H" & Right(QWM, 2)) 'λ��
        
        Dim X, Y, Z, M

'++++++++++++++++++++++++++++++++++++++++++++++++++
'XYZM��ԭʼ���ʽ
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
        For J = 1 To 32 'ÿ����Ϊ32���ֽ�
            ZMHP(i, J) = HZK166(Address + J - 1)  '���������ݴ���,����
            'Debug.Print ZMHP(i, J)
        Next
    Next
End Function
  
'=============================
'�������ȡģ������ZMHP����,����Shape1 ��ʾ
'
'
Function PrintHData() '��ӡHData
    Dim S As String '���ԭ��
    Dim I2 As Integer, i As Integer, J As Integer   'for������
    Dim tmpStr As String  '��ʱ�ַ� ���zmhp(j)
    Dim K As Integer
    K = IIf(HScroll1.Value = 1, 1, N - HScroll1.Value + 2) '??Ϊʲô+2?
    For I2 = K To N
        If Style Then
            S = S & "DB    "
        Else
            S = S & "{"
        End If
        For J = 1 To 32
            
            tmpStr = Hex(ZMHP(I2, J))
            tmpStr = Right("00" & tmpStr, 2) '��ʽ�����ַ�,���Ȳ�Ϊ2��0��.
            '==================����ַ�����ʾ���ַ�
            
            '��ʾshape1
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

Function H2VH()  '������ת��Ϊ����
    Dim i, K, M, M2 As Integer
    Dim Z, ZA As Byte '��ʱ����
    Dim BitI As Byte  '����ȡģת����ȡģ�ĵ�Iλ, I= 8~1
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
                    If ((ZMHP(i, K + M2 * 2) And BitI) <> 0) Then Z = Z + tmpM22     '��λ���λ
                    If ((ZMHP(i, K + M2 * 2 + 16) And BitI) <> 0) Then ZA = ZA + tmpM22
                Next
                ZMVH(i, ((K - 1) * 8) + M) = Z  'ȡ��Ϊ�ϲ���
                ZMVH(i, ((K - 1) * 8) + M + 16) = ZA '��һ����
            Next
        Next
    Next
End Function

Function PrintVH() '��ӡ����ת��Ϊ����
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
            tmpStr = Right("00" & tmpStr, 2) '��ʽ��,���ȱ�Ϊ2 ����ǰ����0��
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

Private Sub HScroll1_Change() 'Ԥ������ڵ�ˮƽ������
    HScroll1_Scroll
End Sub

Private Sub HScroll1_GotFocus() 'Ԥ������ڵ�ˮƽ������
    Text1.SetFocus
End Sub

Private Sub HScroll1_Scroll()
      On Error Resume Next
      Dim I2, J, i As Integer 'for����
      Dim tmpStr As Byte ' As String  '��ʱ�ַ� ���zmhp(j)
      For I2 = 0 To 3
            For J = 1 To 32
                tmpStr = ZMHP(I2 + HScroll1.Value, J)
                '��ʾshape1
                For i = 1 To 8
                    Shape1((J - 1) * 8 + i - 1 + (I2) * 256).BackColor = IIf((tmpStr And 2 ^ (8 - i)) = 0, GetColor(vbRed), GetColor(vbYellow))
                Next
            Next
      Next
      
End Sub


Private Sub Text1_Click()
    'text1 �޽���ʱ���֮��ȫѡ
    If Text1Focus = False Then
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1)
    End If
    Text1Focus = True '���ý���
End Sub

Private Sub command2_Click(Index As Integer)
    
    Clipboard.Clear   ' ��������塣
    Clipboard.SetText Text2(Index).Text  ' �����ķ����ڼ������ϡ�
    
End Sub

Function GetColor(Color) '��ɫ����
    GetColor = Color
End Function

Function GetNeiMa() '��ȡ��������
    Dim i As Integer
    Dim Line1 As String
    Dim line2 As String
    Dim TmpHex As String
    Line1 = "DB    "
    line2 = ";     "
    For i = 1 To UBound(HzStr1)
        TmpHex = Right("0000" & Hex(Asc(HzStr1(i))), 4)
        Line1 = Line1 & "0" & Left(TmpHex, 2) & "H,"
        Line1 = Line1 & "0" & Right(TmpHex, 2) & "H" & IIf(i = UBound(HzStr1), " ;���������(��λ��ǰ,��λ�ں�)", ",")
        line2 = line2 & HzStr1(i) & IIf(i = UBound(HzStr1), "        ;�����ַ�", "        ")
    Next
    Text2(0) = Line1 & vbCrLf & line2
End Function


Private Sub Text1_LostFocus() 'text1 ʧȥ����
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
'1. ���Ӻ�������
'2. ���������Ϊ��������

'Update 2009-08-08
'Loveyu
'1.��������ͼ��
'2.��̬ˢ�¸�Frameλ��.
'3.��ʼ����ֻ����������Ϣ

'Updatae 2010 - 3 - 6!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Loveyu
'1.���ֿ����ϵ���Դ�ļ���,����Ҫ���ⲿ��ȡ�ֿ��ļ�!!!!!!
'2.���ݵ���ͼ���������ɱ���!!!!!!
'3.ûʱ���Ż�����.����BUG
'4.�汾����3.0
