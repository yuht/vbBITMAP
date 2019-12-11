VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   10530
   ClientLeft      =   6960
   ClientTop       =   3600
   ClientWidth     =   17970
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10530
   ScaleMode       =   0  'User
   ScaleWidth      =   17970
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   1905
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "Form1.frx":030A
      Top             =   6480
      Width           =   8610
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   7515
      Top             =   270
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   1800
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "Form1.frx":0310
      Top             =   4635
      Width           =   8610
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   1935
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Form1.frx":0316
      Top             =   2655
      Width           =   8610
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
      Appearance      =   0  'Flat
      Height          =   1845
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form1.frx":031C
      Top             =   765
      Width           =   8610
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Text            =   "�ں���"
      Top             =   60
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000009&
      BackStyle       =   1  'Opaque
      DrawMode        =   3  'Not Merge Pen
      Height          =   135
      Index           =   999
      Left            =   8910
      Shape           =   1  'Square
      Top             =   945
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
  DrawShape
  Chinese2HData
  PrintHData
  H2V
  PrintV
  Me.Cls
End Sub



Private Sub Form_Load()
  '**************************************
  ZitiPath = App.Path & "\" & "hzk16" '
  Text2.Text = "����ԭ����"
  Text3.Text = "���������"
  Command1.Caption = "������ģ"
  E = 1
  Text4.Text = "���������" & vbCrLf & vbCrLf & "�������õ����������"
  Text5.Text = "����ȡ��������"
  

 ' Command1_Click
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
  Dim I2, I, J As Integer
  Dim SI
  
  For I2 = 0 To N - 1
    For I = 0 To 15 '16��
      For J = 0 To 15 '16��
        SI = I2 * 256 + I * 16 + J 'ÿ��16��
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
  Dim I2, I, J As Integer 'for������
  Dim tmpstr As String  '��ʱ�ַ� ���zmhp(j)
  
  For I2 = 1 To N

    For J = 1 To 32
      tmpstr = Hex(ZMHP(I2, J))
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
          Shape1((J - 1) * 8 + I - 1 + (I2 - 1) * 256).BackColor = vbRed
        Else
          Shape1((J - 1) * 8 + I - 1 + (I2 - 1) * 256).BackColor = vbYellow
        End If
      Next
  
      
      S = S & "0x" & tmpstr & ","
  
      tmpstr = Hex(&HFF - ZMHP(I2, J))
      If Len(tmpstr) = 1 Then tmpstr = "0" & tmpstr
      W = W & "0x" & tmpstr & ","
  
      If J Mod 16 = 0 Then
        S = S & vbCrLf
        W = W & vbCrLf
      End If
    Next
            S = S & vbCrLf
        W = W & vbCrLf

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

Function H2V()  '������ת��Ϊ����
  Dim I, K, M, M2 As Integer
  Dim Z, ZA As Byte '��ʱ����
  Dim BitI As Byte  '����ȡģת����ȡģ�ĵ�Iλ, I= 8~1
  ReDim ZMSP(1 To N, 1 To 32)
  Dim tmpM22 As Byte '2^
  
  For I = 1 To N
    For K = 1 To 2 '
      For M = 1 To 8
        Z = 0
        ZA = 0
        BitI = 2 ^ (8 - M)
        For M2 = 0 To 7
          tmpM22 = 2 ^ (7 - M2)
          If ((ZMHP(I, K + M2 * 2) And BitI) <> 0) Then Z = Z + tmpM22     '��λ���λ
          If ((ZMHP(I, K + M2 * 2 + 16) And BitI) <> 0) Then ZA = ZA + tmpM22
        Next
        ZMSP(I, ((K - 1) * 16) + M * 2 - 1) = Z 'ȡ��Ϊ�ϲ���
        ZMSP(I, ((K - 1) * 16) + M * 2) = ZA  '��һ����
      Next
    Next
  Next
End Function

Function PrintV()
  Dim I2, I As Integer
  Dim S, W As String
  Dim tmpstr As String
  
  For I2 = 1 To N
    For I = 1 To 32
    tmpstr = Hex(ZMSP(I2, I))
    If Len(tmpstr) = 1 Then tmpstr = "0" & tmpstr '��ʽ��,���ȱ�Ϊ2 ����ǰ����0��
    S = S & "0x" & tmpstr & ","
    
    tmpstr = Hex(&HFF - ZMSP(I2, I))
    If Len(tmpstr) = 1 Then tmpstr = "0" & tmpstr '��ʽ��,���ȱ�Ϊ2 ����ǰ����0��
    W = W & "0x" & tmpstr & ","  'Debug.Print Hex(&HFF - zmhp(1, i)) & ",";
    If I Mod 16 = 0 Then
      S = S & vbCrLf
      W = W & vbCrLf
    End If
    Next
            S = S & vbCrLf
        W = W & vbCrLf
  Next

  Text4.Text = "����ȡԭ���ݣ�" & vbCrLf & S
  Text5.Text = "����ȡ�������ݣ�" & vbCrLf & W
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


