VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4155
   ClientLeft      =   6960
   ClientTop       =   3600
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4155
   ScaleMode       =   0  'User
   ScaleWidth      =   10545
   Begin VB.TextBox Text4 
      Height          =   1095
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   3000
      Width           =   10395
   End
   Begin VB.TextBox Text3 
      Height          =   1095
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   720
      Width           =   10395
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Default         =   -1  'True
      Height          =   555
      Left            =   6300
      TabIndex        =   0
      Top             =   60
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   1095
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1860
      Width           =   10395
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
      Height          =   555
      Left            =   60
      TabIndex        =   1
      Text            =   "��"
      Top             =   60
      Width           =   5955
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim HZStr1() As String '���ı���ȡ�������Ĵ�ŵ�һ������
Dim N As Integer '�ı�������hzLen()
Dim Address1  '�����hzk16�еĵ�ַ
Dim ZMHP() As Byte '��������ִ�ŵ�����
Dim ZMSP() As Byte '���������ִ�ŵ�����
Dim ZiTiPath '��������·��
Dim E
 'Public LoginSucceeded As Boolean


Private Sub Command1_Click()
'On Error Resume Next
Call S0
Call S1
Call S2
Call S3
Call S4
Me.Cls
End Sub


Private Sub Form_Load()
'**************************************

ZiTiPath = App.Path & "\" & "hzk16" '
'Me.Caption = "write for:������������"

Text2.Text = "����ȡ��"
Text3.Text = "���������"
Text4.Text = "���������" & vbCrLf & vbCrLf & "�������õ����������"

Command1.Caption = "����"

E = 1


'If (LoginSucceeded = True) Then Call Command1_Click


End Sub


Public Sub S0()
  Dim UBoundHZSTR1
  UBoundHZSTR1 = 1
'  N = 0
'  Dim L, I, J As Integer
'  L = Len(Text1.Text)
'
'  ReDim Preserve HZStr1(UBound(HZStr1) + 1)
'
'  For I = 1 To L
'    If Asc(Mid(Text1.Text, I, 1)) < 0 Then
'      N = N + 1 '�õ������ֵ�����
'    End If
'  Next
'

'
'  J = 1
  
    For I = 1 To Len(Text1)
      If Asc(Mid(Text1.Text, I, 1)) < 0 Then
        ReDim Preserve HZStr1(UBoundHZSTR1)
        HZStr1(UBoundHZSTR1) = Mid(Text1.Text, I, 1) '�Ѻ��ִ���������.
        UBoundHZSTR1 = UBound(HZStr1) + 1
'      J = J + 1 'ascii
      End If
  Next
    'N = UBoundHZSTR1
    

  
End Sub

Public Sub S1()
  Dim HZK166() As Byte
  Dim QWM, QM, WM   '
  Dim I As Integer, J As Integer
  'Dim k As Integer
  'Dim bytethz As Byte
  Dim IntFileNum
  'Dim mypath
  Dim FileNa
  'mypath = App.Path & "\"
  FileNa = ZiTiPath
  'filena = mypath & "hzk16"
  IntFileNum = FreeFile
  Open FileNa For Binary As #IntFileNum '
  Sum = LOF(IntFileNum) '
  ReDim HZK166(1 To Sum) '
  Get #IntFileNum, , HZK166 '
  Close #IntFileNum '�ر��ֿ��ļ�,��ֹ��������
  ReDim ZMHP(1 To N, 1 To 32) 'As Byte
  For I = 1 To UBound(HZStr1)
    QWM = Hex(Asc(HZStr1(I)) - &HA0A0)
    If Len(QWM) = 3 Then
      QM = Mid(QWM, 1, 1)
      WM = Mid(QWM, 2, 2)
    ElseIf Len(QWM) = 4 Then
      QM = Mid(QWM, 1, 2)
      WM = Mid(QWM, 3, 2)
    End If
    Address1 = 32 * ((CLng("&H" & QM) - 1) * 94 + (CLng("&H" & WM) - 1))
    For J = 1 To 32 'ÿ����Ϊ32���ֽ�
      'bytehz = Hex(hzk166(address1 + j))
      'If Len(bithz) = 1 Then
      'bithz = 0 & bithz
      'End If
      ZMHP(I, J) = HZK166(Address1 + J) '���������ݴ���,����
      On Error Resume Next
    Next
  Next
End Sub

Public Sub S3() '������ת��Ϊ����
  Dim I As Integer, k As Integer
  Dim J As Integer
  Dim m As Integer
  Dim bithz1 As Byte '�����жϸ�λ��ֵ
  Dim z As Byte '
  Dim qq As Byte '
  ReDim ZMSP(1 To N, 1 To 32)

  For I = 1 To N
    'i = 1
    J = 1
    For k = 1 To 2 '
      'Debug.Print "_____"
      qq = &H80 '
      'If (qq >= &H1) Then
      For m = 1 To 8
      'Debug.Print "_____"'�������16*16������������ŵĴ�����Խ��м���.
        z = &H0
        If ((ZMHP(I, k) And qq) = 0) Then bithz1 = &H0 Else bithz1 = &H1 '8
        z = &H80 * bithz1 '��λ���λ
        If ((ZMHP(I, k + 2) And qq) = 0) Then bithz1 = &H0 Else bithz1 = &H1 '7
        z = z + (&H40 * bithz1)
        If ((ZMHP(I, k + 4) And qq) = 0) Then bithz1 = &H0 Else bithz1 = &H1 '6
        z = z + (&H20 * bithz1)
        If ((ZMHP(I, k + 6) And qq) = 0) Then bithz1 = &H0 Else bithz1 = &H1 '5
        z = z + (&H10 * bithz1)
        If ((ZMHP(I, k + 8) And qq) = 0) Then bithz1 = &H0 Else bithz1 = &H1 '4
        z = z + (&H8 * bithz1)
        If ((ZMHP(I, k + 10) And qq) = 0) Then bithz1 = &H0 Else bithz1 = &H1 '3
        z = z + (&H4 * bithz1)
        If ((ZMHP(I, k + 12) And qq) = 0) Then bithz1 = &H0 Else bithz1 = &H1 '2
        z = z + (&H2 * bithz1)
        If ((ZMHP(I, k + 14) And qq) = 0) Then bithz1 = &H0 Else bithz1 = &H1 '1
        z = z + (&H1 * bithz1) '��Ϊ���λ
        ZMSP(I, J) = z 'ȡ��Ϊ�ϲ���
        J = J + 1
        z = 0
        If ((ZMHP(I, k + 16) And qq) = 0) Then bithz1 = &H0 Else bithz1 = &H1 '8
        z = &H80 * bithz1
        If ((ZMHP(I, k + 18) And qq) = 0) Then bithz1 = &H0 Else bithz1 = &H1 '7
        z = z + (&H40 * bithz1)
        If ((ZMHP(I, k + 20) And qq) = 0) Then bithz1 = &H0 Else bithz1 = &H1 '6
        z = z + (&H20 * bithz1)
        If ((ZMHP(I, k + 22) And qq) = 0) Then bithz1 = &H0 Else bithz1 = &H1 '5
        z = z + (&H10 * bithz1)
        If ((ZMHP(I, k + 24) And qq) = 0) Then bithz1 = &H0 Else bithz1 = &H1 '4
        z = z + (&H8 * bithz1)
        If ((ZMHP(I, k + 26) And qq) = 0) Then bithz1 = &H0 Else bithz1 = &H1 '3
        z = z + (&H4 * bithz1)
        If ((ZMHP(I, k + 28) And qq) = 0) Then bithz1 = &H0 Else bithz1 = &H1 '2
        z = z + (&H2 * bithz1)
        If ((ZMHP(I, k + 30) And qq) = 0) Then bithz1 = &H0 Else bithz1 = &H1 '1
        z = z + (&H1 * bithz1)
        ZMSP(I, J) = z '��һ����
        J = J + 1 '            ��һ������Щ�鷳,Ҫ���Ļ���ע�����vb��16������������
        z = 0
        qq = (qq / &H2) 'ȡ���ŵ���һλ*****************��������Ը����ܿ�������������������������������
      Next
    Next
  Next
End Sub '�����Ŵ�ŵĺ���ת��λ,���Ŵ�ŵ�����vbû����λ����ֻ�а�λ����,�����ж�.

Public Sub S2()
  Dim s As String
  Dim w As String
  For I = 1 To 8
    'Debug.Print Hex(&HFF - zmhp(1, i)) & ",";
    s = s & "0" & Hex((ZMHP(1, I))) & ","
    w = w & "0" & Hex((&HFF - ZMHP(1, I))) & "H,"
  
  Next I
  's = s & vbCrLf
  'w = w & vbCrLf
  For I = 9 To 16
    'Debug.Print Hex(&HFF - zmhp(1, i)) & ",";
    s = s & "0" & Hex((ZMHP(1, I))) & ","
    w = w & "0" & Hex((&HFF - ZMHP(1, I))) & "H,"
  Next I
  s = s & vbCrLf
  w = w & vbCrLf
  For I = 17 To 24
    'Debug.Print Hex(&HFF - zmhp(1, i)) & ",";
    s = s & "0" & Hex((ZMHP(1, I))) & ","
    w = w & "0" & Hex((&HFF - ZMHP(1, I))) & "H,"
  Next I
  's = s & vbCrLf
  'w = w & vbCrLf
  For I = 25 To 32
    'Debug.Print Hex(&HFF - zmhp(1, i)) & ",";
    s = s & "0" & Hex((ZMHP(1, I))) & ","
    w = w & "0" & Hex((&HFF - ZMHP(1, I))) & "H,"
  Next I
  s = s & vbCrLf
  w = w & vbCrLf
  
  
  Text3.Text = "����ȡԭ���ݣ�" & vbCrLf & s
  Text2.Text = "����ȡ�������ݣ�" & vbCrLf & w
End Sub

Public Sub S4()
  Dim s As String
  Dim w As String
  For I = 1 To 8
    'Debug.Print Hex(&HFF - zmhp(1, i)) & ",";
    s = s & "0" & Hex((ZMSP(1, I))) & "H,"
  Next
  's = s & vbCrLf
  For I = 9 To 16
    'Debug.Print Hex(&HFF - zmhp(1, i)) & ",";
    s = s & "0" & Hex((ZMSP(1, I))) & "H,"
  Next
  
  s = s & vbCrLf
  
  For I = 17 To 24
    'Debug.Print Hex(&HFF - zmhp(1, i)) & ",";
    s = s & "0" & Hex((ZMSP(1, I))) & "H,"
  Next I
  
  's = s & vbCrLf
  
  For I = 25 To 32
    'Debug.Print Hex(&HFF - zmhp(1, i)) & ",";
    s = s & "0" & Hex((ZMSP(1, I))) & "H,"
  Next I
  
  s = s & vbCrLf

  Text4.Text = "����ȡԭ���ݣ�" & vbCrLf & s
  'Text2.Text = "����ȡ�������ݣ�" & vbCrLf & w
End Sub

'��������-2007-12-16 QQ:102126913 tel:
'write for ��͸˲��
'*******************************************************
'rewrite 2008-1-8
'write for snail boy
'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Label1.Caption = "(x,y)=" & "(" & X & "," & Y & ")"
'Me.Line (X, Y)-(X + 100, Y + 100), vbRed
'End Sub
