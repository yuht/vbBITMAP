VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "������ģ"
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
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame Frame3 
      Caption         =   "��ģ���"
      Height          =   5610
      Left            =   45
      TabIndex        =   9
      Top             =   2835
      Width           =   9315
      Begin VB.CommandButton Command3 
         Caption         =   "Command3 ���Ƶ�������"
         Height          =   375
         Left            =   7575
         TabIndex        =   5
         Top             =   2880
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2 ���Ƶ�������"
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
      Caption         =   "���뺺��"
      Height          =   750
      Left            =   45
      TabIndex        =   8
      Top             =   45
      Width           =   9315
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
         Left            =   180
         TabIndex        =   0
         Top             =   240
         Width           =   7455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1 ������ģ"
         Default         =   -1  'True
         Height          =   375
         Left            =   7740
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "����Ԥ��"
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
Dim HzStr1() As String '���ı���ȡ�������Ĵ�ŵ�һ������
Dim ZitiPath   As String '��������·��

Dim N As Integer '�ı�������hzLen()
Dim Address1  '�����hzk16�еĵ�ַ
Dim ZMHP(), ZMSP(), ZMVH() As Byte '�����/�����ִ�ŵ�����

  Dim SWidth As Integer
  Dim Spre As Integer
  Dim SPre8 As Integer
  Dim SPreChr As Integer


Private Sub Command1_Click()
  Me.Cls
  SaveChinese '�� text1���ҳ�����
  
  If N = 0 Then
    Text1 = "�ں���"
    Text1.SetFocus
    Command1_Click 'Exit Sub
  End If
  DrawShape '����ģ�µ���
  Chinese2HData '����˳��ȡģ ����ZMHP��
  PrintHData  '���ZMHP JD240128-7
  H2VH '���ֺ�Դ���� JD12864C-1 Һ��������
  PrintVH '��ӡZMVH

End Sub





Private Sub Form_Load()
  
  On Error Resume Next
  ZitiPath = App.Path & "\" & "hzk16" '
  Text2.Text = ""
  Text3.Text = ""
  Command1.Caption = "������ģ"
  Command2.Caption = "���Ƶ�������"
  Command3.Caption = "���Ƶ�������"
'  Label1 = "�������� ����ȡԭ����" & vbCrLf & ""
 Label1 = vbCrLf & "(���ֺ�Դ����  JD240128-7 Һ����ʾ��ר��):"
'  Label2 = "��λ��ǰ,��������,��������" & vbCrLf & "(���ֺ�Դ���� JD12864C-1 Һ����ʾ��ר��):"
 Label2 = vbCrLf & "(���ֺ�Դ����  JD12864C-1 Һ����ʾ��ר��):"

  If Len(Command) <> 0 Then
    Text1 = Command
    Call Command1_Click
  End If
  
End Sub

'=======================
'��Text1���ҳ�����,��ŵ�HzStr1������
'���ָ�����ŵ� N ��
Function SaveChinese()
  ReDim HzStr1(0)
  Dim UboundStr1, I
  For I = 1 To Len(Text1)
    If Asc(Mid(Text1.Text, I, 1)) < 0 Then
      UboundStr1 = UBound(HzStr1) + 1
      ReDim Preserve HzStr1(UboundStr1)
      HzStr1(UboundStr1) = Mid(Text1.Text, I, 1) '�Ѻ��ִ���������.
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
        For I = 0 To 15 '16��
            For J = 0 To 15 '16��
                SI = I2 * 256 + I * 16 + J   'ÿ��16��
                
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
' ����˳�����ȡģ,��ŵ�ZMHP������
' ����ģ�ļ���ŵ�Hzk166��
' �ҵ�������ģλ��,������˳���ŵ�ZMHP(���ָ���,32)������
'
Function Chinese2HData()
      Dim Hzk166() As Byte '�ֿ�����
      Dim Sum, QWM, QM, WM  '�ļ�����,��λ��,����,λ��
      Dim intFileNum  '�ļ���
      Dim FileNa  '�ļ���
      Dim I, J As Integer
      
    '===============
    '��ȡ�ֿ⵽����
      
      FileNa = ZitiPath '��ģ·��
      intFileNum = FreeFile '�����ļ���
      Open FileNa For Binary As #intFileNum   '���ļ�
      Sum = LOF(intFileNum)   '�ļ�����
          
      If Sum = 0 Then
            Close #intFileNum
            Kill FileNa
            MsgBox "�뽫���ֿ��ļ�hzk16(748 KB (766,080 �ֽ�) �� 261 KB (267,616 �ֽ�)) ���뵱ǰĿ¼��!", vbExclamation
            End
      End If
      
      
      ReDim Hzk166(1 To Sum)  '���ļ����ȶ�������
      Get #intFileNum, , Hzk166 '���뵽���� Hzk166
      Close #intFileNum '�ر��ֿ��ļ�,��ֹ��������
      
      ReDim ZMHP(1 To N, 1 To 32) '�������� Nά
    
      '===================
      '��ȡ������λ��
      For I = 1 To UBound(HzStr1)
            QWM = Hex(Asc(HzStr1(I)) - &HA0A0) '��λ��
            QWM = Right("0000" & QWM, 4)
            QM = Left(QWM, 2) '����
            WM = Right(QWM, 2) 'λ��
        
            '================
            '��ȡ���ֵ�����ʼλ��
            Address1 = 32 * ((CLng("&H" & QM) - 1) * 94 + (CLng("&H" & WM) - 1))
            For J = 1 To 32 'ÿ����Ϊ32���ֽ�
                ZMHP(I, J) = Hzk166(Address1 + J)    '���������ݴ���,����
            Next
      Next
End Function
  
'=============================
'�������ȡģ������ZMHP����,����Shape1 ��ʾ
'
'
Function PrintHData()
    Dim S As String '���ԭ��
    Dim I2, I, J As Integer 'for������
    Dim tmpStr As String  '��ʱ�ַ� ���zmhp(j)
    
    For I2 = 1 To N
        S = S & "DB    "
        For J = 1 To 32
            tmpStr = Hex(ZMHP(I2, J))
            tmpStr = Right("00" & tmpStr, 2) '��ʽ�����ַ�,���Ȳ�Ϊ2��0��.
            '==================����ַ�����ʾ���ַ�
            
            '��ʾshape1
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

Function H2VH()  '������ת��Ϊ����
    Dim I, K, M, M2 As Integer
    Dim Z, ZA As Byte '��ʱ����
    Dim BitI As Byte  '����ȡģת����ȡģ�ĵ�Iλ, I= 8~1
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
                    If ((ZMHP(I, K + M2 * 2) And BitI) <> 0) Then Z = Z + tmpM22     '��λ���λ
                    If ((ZMHP(I, K + M2 * 2 + 16) And BitI) <> 0) Then ZA = ZA + tmpM22
                Next
                ZMVH(I, ((K - 1) * 8) + M) = Z  'ȡ��Ϊ�ϲ���
                ZMVH(I, ((K - 1) * 8) + M + 16) = ZA '��һ����
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
            tmpStr = Right("00" & tmpStr, 2) '��ʽ��,���ȱ�Ϊ2 ����ǰ����0��
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
      Dim I2, J, I As Integer 'for����
      Dim tmpStr As String  '��ʱ�ַ� ���zmhp(j)
      For I2 = 0 To 6
            For J = 1 To 32
                tmpStr = Hex(ZMHP(I2 + HScroll1.Value, J))
                '��ʾshape1
                For I = 1 To 8
                    Shape1((J - 1) * 8 + I - 1 + (I2) * 256).BackColor = IIf((Val("&h" & tmpStr) And 2 ^ (8 - I)) = 0, GetColor(vbRed), GetColor(vbYellow))
                Next
            Next
      Next
End Sub
 


Private Sub Text1_Click()
    '���֮��ȫѡ
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)
End Sub

Private Sub command2_Click()
    
    Clipboard.Clear   ' ��������塣
    Clipboard.SetText Text2.Text  ' �����ķ����ڼ������ϡ�
End Sub

Private Sub command3_Click()
    Clipboard.Clear   ' ��������塣
    Clipboard.SetText Text3.Text  ' �����ķ����ڼ������ϡ�
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

