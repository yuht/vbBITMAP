VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4710
   ClientLeft      =   6975
   ClientTop       =   3615
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleMode       =   0  'User
   ScaleWidth      =   6660
   Begin VB.TextBox Text4 
      Height          =   1215
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "Form1.frx":0000
      Top             =   3420
      Width           =   6495
   End
   Begin VB.TextBox Text3 
      Height          =   1215
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Form1.frx":0006
      Top             =   780
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
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form1.frx":000C
      Top             =   2100
      Width           =   6495
   End
   Begin VB.TextBox Text1 
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
      Width           =   615
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
  SaveChinese
  Chinese2Data
  Data2H
  H2V
  PrintV
  Me.Cls
End Sub



Private Sub Form_Load()
  '**************************************
  ZitiPath = App.Path & "\" & "hzk16" '
  Text2.Text = "����ȡ��"
  'Me.Caption = "write for:������������"
  Text3.Text = "���������"
  Command1.Caption = "������ģ"
  E = 1
  Text4.Text = "���������" & vbCrLf & vbCrLf & "�������õ����������"
End Sub


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

Function Chinese2Data()
  Dim Hzk166() As Byte '�ֿ�����
  Dim Sum, QWM, QM, WM  '�ļ�����,��λ��,����,λ��
  Dim I, J As Integer
  Dim intFileNum  '�ļ���
  Dim FileNa  '�ļ���
'===============
'��ȡ�ֿ⵽����
  
  FileNa = ZitiPath
  intFileNum = FreeFile
  Open FileNa For Binary As #intFileNum   '���ļ�
  Sum = LOF(intFileNum)   '�ļ�����
  ReDim Hzk166(1 To Sum)  '���ļ����ȶ�������
  Get #intFileNum, , Hzk166 '��������
  Close #intFileNum '�ر��ֿ��ļ�,��ֹ��������
  
  ReDim ZMHP(1 To N, 1 To 32) 'As Byte

  '===================
  '��ȡ������λ��
  For I = 1 To UBound(HzStr1)
    QWM = Hex(Asc(HzStr1(I)) - &HA0A0)
    If Len(QWM) = 3 Then
      QM = Left(QWM, 1)
    ElseIf Len(QWM) = 4 Then
      QM = Left(QWM, 2)
    End If
    WM = Right(QWM, 2)

    '================
    '��ȡ���ֵ���
    Address1 = 32 * ((CLng("&H" & QM) - 1) * 94 + (CLng("&H" & WM) - 1))
    For J = 1 To 32 'ÿ����Ϊ32���ֽ�
      ZMHP(I, J) = Hzk166(Address1 + J)    '���������ݴ���,����
'      On Error Resume Next
    Next
  Next
  End Function
  
Function Data2H()
  Dim S, W As String
  Dim I, J As Integer
  Dim StartN, EndN As Integer
  Dim TmpStr As String
  For I = 1 To 4
    EndN = I * 8
    StartN = EndN - 7
    
    For J = StartN To EndN
      TmpStr = Hex(ZMHP(1, J))
      If Len(TmpStr) = 1 Then TmpStr = "0" & TmpStr
      S = S & "0x" & TmpStr & ","
      
      TmpStr = Hex(&HFF - ZMHP(1, J))
      If Len(TmpStr) = 1 Then TmpStr = "0" & TmpStr
      W = W & "0x" & TmpStr & ","
    Next
    S = S & vbCrLf
    W = W & vbCrLf
  Next

  Text3 = "����ȡԭ���ݣ�" & vbCrLf & S
  Text2 = "����ȡ�������ݣ�" & vbCrLf & W

End Function

'============================================
'��ZMHP �����ַ�����תΪ��������ZMSP
'��ZMHP��1,3,5,7~~~~ 15 �ĵ�8λ��ΪZMSP�ĵ�1��
'��ZMHP��17,16,21,23~~~~ 32 �ĵ�8λ��ΪZMSP�ĵ�2��
'
'���ؽ����ŵ�ZMSP
'
Function H2V() '������ת��Ϊ����
  Dim I, K, J, M As Integer
  Dim BitHZ1 As Byte '�����жϸ�λ��ֵ
  Dim Z As Byte
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
        Z = &H0
        
        If ((ZMHP(I, K) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '8
        Z = &H80 * BitHZ1 '��λ���λ
        
        If ((ZMHP(I, K + 2) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '7
        Z = Z + (&H40 * BitHZ1)
        
        If ((ZMHP(I, K + 4) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '6
        Z = Z + (&H20 * BitHZ1)
        
        If ((ZMHP(I, K + 6) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '5
        Z = Z + (&H10 * BitHZ1)
        
        If ((ZMHP(I, K + 8) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '4
        Z = Z + (&H8 * BitHZ1)
        
        If ((ZMHP(I, K + 10) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '3
        Z = Z + (&H4 * BitHZ1)
        
        If ((ZMHP(I, K + 12) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '2
        Z = Z + (&H2 * BitHZ1)
        
        If ((ZMHP(I, K + 14) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '1
        Z = Z + (&H1 * BitHZ1) '��Ϊ���λ
        
        ZMSP(I, J) = Z 'ȡ��Ϊ�ϲ���
        J = J + 1
        Z = 0
        
        If ((ZMHP(I, K + 16) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '8
        Z = &H80 * BitHZ1
        
        If ((ZMHP(I, K + 18) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '7
        Z = Z + (&H40 * BitHZ1)
        
        If ((ZMHP(I, K + 20) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '6
        Z = Z + (&H20 * BitHZ1)
        
        If ((ZMHP(I, K + 22) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '5
        Z = Z + (&H10 * BitHZ1)
        
        If ((ZMHP(I, K + 24) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '4
        Z = Z + (&H8 * BitHZ1)
        
        If ((ZMHP(I, K + 26) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '3
        Z = Z + (&H4 * BitHZ1)
        
        If ((ZMHP(I, K + 28) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '2
        Z = Z + (&H2 * BitHZ1)
        
        If ((ZMHP(I, K + 30) And BitI) = 0) Then BitHZ1 = &H0 Else BitHZ1 = &H1 '1
        Z = Z + (&H1 * BitHZ1)
        
        ZMSP(I, J) = Z '��һ����
        J = J + 1 '            ��һ������Щ�鷳,Ҫ���Ļ���ע�����vb��16������������
        Z = 0
        
        BitI = (BitI / &H2) 'ȡ���ŵ���һλ*****************��������Ը����ܿ�������������������������������
      Next
    Next
  Next
End Function '�����Ŵ�ŵĺ���ת��λ,���Ŵ�ŵ�����vbû����λ����ֻ�а�λ����,�����ж�.

Function PrintV()
  Dim I  As Integer
  Dim S, W As String
  For I = 1 To 8
  'Debug.Print Hex(&HFF - zmhp(1, i)) & ",";
  S = S & "0x" & Hex((ZMSP(1, I))) & ","
  Next
  
  S = S & vbCrLf
  
  For I = 9 To 16
  'Debug.Print Hex(&HFF - zmhp(1, i)) & ",";
  S = S & "0x" & Hex((ZMSP(1, I))) & ","
  Next
  
  S = S & vbCrLf
  
  For I = 17 To 24
  'Debug.Print Hex(&HFF - zmhp(1, i)) & ",";
  S = S & "0x" & Hex((ZMSP(1, I))) & ","
  Next I
  
  S = S & vbCrLf
  
  For I = 25 To 32
  'Debug.Print Hex(&HFF - zmhp(1, i)) & ",";
  S = S & "0x" & Hex((ZMSP(1, I))) & ","
  Next I
  
  S = S & vbCrLf

  Text4.Text = "����ȡԭ���ݣ�" & vbCrLf & S
  'Text2.Text = "����ȡ�������ݣ�" & vbCrLf & w

End Function

'��������-2007-12-16 QQ:102126913 tel:
'write for ��͸˲��
'*******************************************************
'rewrite 2008-1-8
'write for snail boy
