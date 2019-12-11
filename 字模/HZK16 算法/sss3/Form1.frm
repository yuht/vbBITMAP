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
         Name            =   "����"
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
      Text            =   "��"
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
Dim HzStr1() As String '���ı���ȡ�������Ĵ�ŵ�һ������
Dim ZitiPath   As String '��������·��

Dim N As Integer '�ı�������hzLen()
Dim Address1  '�����hzk16�еĵ�ַ
Dim ZMHP(), ZMSP() As Byte '�����/�����ִ�ŵ�����
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
  Text2.Text = "����ȡ��"
  Text3.Text = "���������"
  Command1.Caption = "������ģ"
  E = 1
  Text4.Text = "���������" & vbCrLf & vbCrLf & "�������õ����������"
  
  '=================================
  'Shape1 ��ɵ�16*16����
  On Error Resume Next
  Dim I, J, SI As Integer
  For I = 0 To 15 '16��
    For J = 0 To 15 '16��
      SI = I * 16 + J   'ÿ��16��
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
  ReDim Hzk166(1 To Sum)  '���ļ����ȶ�������
  Get #intFileNum, , Hzk166 '���뵽���� Hzk166
  Close #intFileNum '�ر��ֿ��ļ�,��ֹ��������
  
  ReDim ZMHP(1 To N, 1 To 32) '�������� Nά

  '===================
  '��ȡ������λ��
  For I = 1 To UBound(HzStr1)
    QWM = Hex(Asc(HzStr1(I)) - &HA0A0) '��λ��
    If Len(QWM) = 3 Then
      QM = Left(QWM, 1) '����
    ElseIf Len(QWM) = 4 Then
      QM = Left(QWM, 2) '����
    End If
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
  Dim S, W As String '���ԭ��,ȡ����
  Dim I, J As Integer 'for������
  Dim tmpstr As String  '��ʱ�ַ� ���zmhp(j)
  

  For J = 1 To 32
    tmpstr = Hex(ZMHP(1, J))
    If Len(tmpstr) = 1 Then tmpstr = "0" & tmpstr '��ʽ�����ַ�,���Ȳ�Ϊ2��0��.
    '==================����ַ�����ʾ���ַ�
        
    '��ʾshape1
    For I = 1 To 8
      '��ʱTimer1
'      Timer1.Enabled = True
'      While Timer1.Enabled = True
'      DoEvents
'      Wend
      '������Ӧλ�������ӦShape1
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

  Text2 = "����ȡԭ���ݣ�" & vbCrLf & S
  Text3 = "����ȡ�������ݣ�" & vbCrLf & W

End Function

'============================================
'��ZMHP �����ַ�������������ZMSP˳������/��������ZMSP
'��ZMHP��1,3,5,7~~~~ 15 �ĵ�8λ��ΪZMSP�ĵ�1��
'��ZMHP��17,16,21,23~~~~ 32 �ĵ�8λ��ΪZMSP�ĵ�2��
'
'���ؽ����ŵ�ZMSP
'
Function H2V() '������ת��Ϊ����
  Dim I, K, J, M As Integer
  Dim BitHZ1 As Byte '�����жϸ�λ��ֵ
  Dim Z, ZA As Byte
  Dim BitI As Byte  '����ȡģת����ȡģ�ĵ�Iλ, I= 8~1
  ReDim ZMSP(1 To N, 1 To 32)
  
  For I = 1 To N
    'i = 1
    J = 1
    For K = 1 To 2 '
      'Debug.Print "_____"
      BitI = &H80 '
      'If (biti >= &H1) Then
      For M = 1 To 8
        'Debug.Print "_____"'�������16*16������������ŵĴ�����Խ��м���.
        Z = 0
        ZA = 0
        
        If ((ZMHP(I, K) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '8
        Z = &H80 * BitHZ1 '��λ���λ
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
        Z = Z + (&H1 * BitHZ1) '��Ϊ���λ
        If ((ZMHP(I, K + 30) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '1
        ZA = ZA + (&H1 * BitHZ1)
        
        ZMSP(I, J) = Z 'ȡ��Ϊ�ϲ���
        J = J + 1
        
        ZMSP(I, J) = ZA '��һ����
        J = J + 1 '            ��һ������Щ�鷳,Ҫ���Ļ���ע�����vb��16������������

        
        BitI = (BitI / &H2) 'ȡ���ŵ���һλ*****************��������Ը����ܿ�������������������������������
      Next
    Next
  Next
End Function '�����Ŵ�ŵĺ���ת��λ,���Ŵ�ŵ�����vbû����λ����ֻ�а�λ����,�����ж�.

Function PrintV()
  Dim I  As Integer
  Dim S, W As String
  Dim tmpstr As String
  
  For I = 1 To 32
  tmpstr = Hex(ZMSP(1, I))
  If Len(tmpstr) = 1 Then tmpstr = "0" & tmpstr '��ʽ��,���ȱ�Ϊ2 ����ǰ����0��
  'Debug.Print Hex(&HFF - zmhp(1, i)) & ",";
  S = S & "0x" & tmpstr & ","
  If I Mod 8 = 0 Then S = S & vbCrLf
  Next


  Text4.Text = "����ȡԭ���ݣ�" & vbCrLf & S
  'Text2.Text = "����ȡ�������ݣ�" & vbCrLf & w

End Function

Function H2V2() '������ת��Ϊ����
  Dim I, K, J, M As Integer
  Dim BitHZ1 As Byte '�����жϸ�λ��ֵ
  Dim Z, ZA As Byte
  Dim BitI As Byte  '����ȡģת����ȡģ�ĵ�Iλ, I= 8~1
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
        'Debug.Print "_____"'�������16*16������������ŵĴ�����Խ��м���.
        Z = 0
        ZA = 0
          tmpM22 = &H80
        For M2 = 0 To 7
        
          If ((ZMHP(I, K + M2 * 2) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '8
          Z = Z + tmpM22 * BitHZ1 '��λ���λ
          If ((ZMHP(I, K + M2 * 2 + 16) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '8
          ZA = ZA + tmpM22 * BitHZ1
                
          tmpM22 = tmpM22 / &H2
        Next
        BitI = BitI / &H2
        
        ZMSP(I, J) = Z 'ȡ��Ϊ�ϲ���
        Debug.Print J
        ZMSP(I, J + 1) = ZA '��һ����
        
        J = J + 2 '            ��һ������Щ�鷳,Ҫ���Ļ���ע�����vb��16������������


      Next
    Next
  Next
End Function '�����Ŵ�ŵĺ���ת��λ,���Ŵ�ŵ�����vbû����λ����ֻ�а�λ����,�����ж�.

Function PrintV2()
  Dim I  As Integer
  Dim S, W As String
  Dim tmpstr As String
  
  For I = 1 To 32
  tmpstr = Hex(ZMSP(1, I))
  If Len(tmpstr) = 1 Then tmpstr = "0" & tmpstr '��ʽ��,���ȱ�Ϊ2 ����ǰ����0��
  'Debug.Print Hex(&HFF - zmhp(1, i)) & ",";
  S = S & "0x" & tmpstr & ","
  If I Mod 8 = 0 Then S = S & vbCrLf
  Next


  Text5.Text = "����ȡԭ����2��" & vbCrLf & S
  'Text2.Text = "����ȡ�������ݣ�" & vbCrLf & w

End Function


Private Sub Timer1_Timer()
  Timer1.Enabled = False
End Sub

'��������-2007-12-16 QQ:102126913 tel:
'write for ��͸˲��
'*******************************************************
'rewrite 2008-1-8
'write for snail boy

'REWRITE 2009-05-16
'WRITE BY LOVEYU
'WRITE FOR MYSELF


