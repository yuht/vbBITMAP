VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "������ģ8*16"
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
         Left            =   7620
         TabIndex        =   5
         Top             =   180
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2 ���Ƶ�������"
         Height          =   375
         Left            =   7620
         TabIndex        =   3
         Top             =   180
         Visible         =   0   'False
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
         Visible         =   0   'False
         Width           =   9150
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   2205
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Text            =   "Form1.frx":0310
         Top             =   630
         Width           =   9150
      End
      Begin VB.Label Label2 
         Caption         =   "Label2(1)"
         Height          =   2550
         Index           =   1
         Left            =   180
         TabIndex        =   12
         Top             =   2970
         Width           =   8940
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   435
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   180
         Width           =   7275
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   435
         Left            =   90
         TabIndex        =   10
         Top             =   180
         Visible         =   0   'False
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
Dim N As Integer '�ı�������hzLen()
Dim ZMHP(), ZMSP(), ZMVH() As Byte '�����/�����ִ�ŵ�����
Dim Hzk166() As Byte '�ֿ�����
Dim SWidth As Integer
Dim Spre As Integer
Dim SPre8 As Integer
Dim SPreChr As Integer

Private Sub Command1_Click()
    Me.Cls
    SaveChinese '�� text1���ҳ�����
    
    If N = 0 Then
        Text1 = ""
        Exit Sub
'        Text1 =     "�ں���"
'        Text1.SetFocus
'        Command1_Click 'Exit Sub
    End If
    DrawShape '����ģ�µ���
    Chinese2HData '����˳��ȡģ ����ZMHP��
    'PrintHData  '���ZMHP JD240128-7
    ReDuleZMHP '˫�ַ��ϲ�Ϊһ������.
    H2VH '���ֺ�Դ���� JD12864C-1 Һ��������
    PrintVH '��ӡZMVH
End Sub


Private Sub Form_Load()
  
    On Error Resume Next
'    Dim ZitiPath   As String '��������·��
'    ZitiPath = App.Path & "\" & "HZK16.GBK" '
'    '===============
'    '��ȡ�ֿ⵽����
'    Dim intFileNum  '�ļ���
'    Dim Sum
'    intFileNum = FreeFile '�����ļ���
'    Open ZitiPath For Binary As #intFileNum    '���ļ�
'    Sum = LOF(intFileNum)   '�ļ�����
'
'    If Sum <> 766080 Then
'        Close #intFileNum
'        Kill ZitiPath
'        Dim tmpFileName As String
'        tmpFileName = "HZK16.GBK (748 KB (766,080 �ֽ�))"
'        MsgBox vbCrLf & "�뽫���ֿ��ļ�" & tmpFileName & "���뵱ǰĿ¼��!" & IIf(Sum <> 0, vbCrLf & "�����ֿ�(��С: " & Sum & " �ֽ� )��ɾ��!", ""), vbExclamation, tmpFileName & "δ�ҵ�!!  - -!"
'        End
'    End If
'
'    ReDim Hzk166(1 To Sum)  '���ļ����ȶ�������
'    Get #intFileNum, , Hzk166 '���뵽���� Hzk166
'    Close #intFileNum '�ر��ֿ��ļ�,��ֹ��������

    Text2.Text = ""
    Text3.Text = ""
    Command1.Caption = "������ģ"
    Command2.Caption = "���Ƶ�������"
    Command3.Caption = "���Ƶ�������"
    'Label1 = "�������� ����ȡԭ����" & vbCrLf & ""
    Label1 = vbCrLf & "(���ֺ�Դ��ѧ�������޹�˾  JD240128-7 Һ��ר��):"
    'Label2 = "��λ��ǰ,��������,��������" & vbCrLf & "(���ֺ�Դ���� JD12864C-1 Һ����ʾ��ר��):"
    Label2(0) = vbCrLf & "(���ֺ�Դ��ѧ�������޹�˾  JD12864C-1 Һ��ר��):"
    Label2(1) = "˵��:" & vbCrLf & "ֻ�ʺϰ��Ӣ��ȡ��ģ,ÿ����ĸռ��8*16, ÿ����һ��Ӣ���ַ�."
    If Len(Command) <> 0 Then
        Text1 = Command
        Call Command1_Click
    End If
  
End Sub

'=======================
'��Text1���ҳ�����,��ŵ�HzStr1������
'���ָ�����ŵ� N ��
Function SaveChinese()
  
    Dim LenText As Integer, I As Integer
    LenText = Len(Text1)
    If LenText Mod 2 <> 0 Then
        LenText = LenText + 1
        ReDim HzStr1(LenText)
        Text1.Text = Text1.Text & " "
    Else
        ReDim HzStr1(LenText)
    End If

    

    
    For I = 1 To LenText
      'If Asc(Mid(Text1.Text, I, 1)) < 0 Then
      '  UboundStr1 = UBound(HzStr1) + 1
      '  ReDim Preserve HzStr1(UboundStr1)
        HzStr1(I) = Mid(Text1.Text, I, 1) '�Ѻ��ִ���������.
      'End If
    Next
    
    N = LenText
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
    
    Dim N1 As Integer
    N1 = N
    N1 = (N1 Mod 2) + Int(N1 / 2)
  
    TK = IIf(N1 > 7, 6, N1 - 1)
    
    
    
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
    HScroll1.Max = N - 14
    HScroll1.Visible = IIf(N > 14, True, False)
End Function


'========================
' ����˳�����ȡģ,��ŵ�ZMHP������
' ����ģ�ļ���ŵ�Hzk166��
' �ҵ�������ģλ��,������˳���ŵ�ZMHP(���ָ���,32)������
'
Function Chinese2HData()
    Dim Address  '�����hzk16�еĵ�ַ
    

    Dim QWM, QM, WM  '�ļ�����,��λ��,����,λ��

    Dim I As Integer, J As Integer
    
    
    ReDim ZMHP(1 To N, 1 To 32) '�������� Nά
    
    '===================
    '��ȡ������λ��
    For I = 1 To UBound(HzStr1)
        If Asc(HzStr1(I)) < 0 Then
            QWM = Hex(Asc(HzStr1(I)))  '��λ��
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
                ZMHP(I, J) = Hzk166(Address + J)    '���������ݴ���,����
            Next
        Else
                
            Dim KK As Integer
            Dim KS As Integer
            Dim DZ
            
            DZ = Split(GetAsc8_16(HzStr1(I)), ",")
            For J = 1 To 16 'ÿ����Ϊ32���ֽ�
                ZMHP(I, J * 2 - 1) = DZ(J - 1) '���������ݴ���,����
            Next
        End If
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
    Dim LengthN As Integer
        
        
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
    ReDuleZMHP
End Function

Function ReDuleZMHP()
    ReDim EngAsc(N, 32) 'Ӣ���ַ�����������
    
    Dim S As String '���ԭ��
    Dim I2, I, J As Integer 'for������
    Dim tmpStr As String  '��ʱ�ַ� ���zmhp(j)
    Dim LengthN As Integer
    
    Dim W1 As Integer, W2 As Integer
    
    '��2��Ӣ��16*32�ַ����һ�� 32*32
    For I2 = 1 To N
        For J = 1 To 32
            W1 = (I2 * 2) - (J Mod 2)
            If W1 > N Then Exit For
            W2 = J - IIf((J Mod 2) = 0, 1, 0)
            EngAsc(I2, J) = ZMHP(W1, W2)
'            Debug.Print I2; J; W1; W2
        Next
    Next
    
    
    '�ַ���������
    Dim KLength As Integer
    Dim CharN As Integer
    Dim sTmp As String
    CharN = 0
    KLength = (N Mod 2) + Int(N / 2)
    N = KLength
    For I2 = 1 To KLength
        S = S & "DB    "
        
        
        
        CharN = CharN + 1
        If (CharN > UBound(HzStr1)) Then
            sTmp = " "
        Else
            sTmp = HzStr1(CharN)
        End If
        CharN = CharN + 1
        If (CharN > UBound(HzStr1)) Then
            sTmp = sTmp & " "
        Else
            sTmp = sTmp & HzStr1(CharN)
        End If


        For J = 1 To 32
            tmpStr = Hex(EngAsc(I2, J))
            ZMHP(I2, J) = EngAsc(I2, J)
            tmpStr = Right("00" & tmpStr, 2) '��ʽ�����ַ�,���Ȳ�Ϊ2��0��.
            '==================����ַ�����ʾ���ַ�
            '��ʾshape1
            For I = 1 To 8
                If I2 > 7 Then Exit For
                Shape1((J - 1) * 8 + I - 1 + (I2 - 1) * 256).BackColor = IIf((Val("&h" & tmpStr) And 2 ^ (8 - I)) = 0, GetColor(vbRed), GetColor(vbYellow))
            Next
            S = S & "0" & tmpStr & "H,"
            If J = 16 Then
                S = S & ";" & sTmp & "-" & I2 & vbCrLf & "DB    "
            End If
        Next

        S = S & ";" & sTmp & "-" & I2 & vbCrLf
        
        
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
    Dim CharN As Integer
    Dim sTmp As String
    Dim I2, I As Integer
    Dim S, W As String
    Dim tmpStr As String
    CharN = 0
    For I2 = 1 To N
        S = S & "DB    "
        
'        For I = 1 To 32
'            tmpStr = Hex(ZMVH(I2, I))
'            tmpStr = Right("00" & tmpStr, 2) '��ʽ��,���ȱ�Ϊ2 ����ǰ����0��
'            S = S & "0" & tmpStr & "H,"
'
'            If I = 16 Then
'                CharN = CharN + 1
'                S = S & ";" & HzStr1(CharN) & "-" & I2 & vbCrLf & "DB    "
'            End If
'
'        Next
        
        
        For I = 1 To 8
            tmpStr = Hex(ZMVH(I2, I))
            tmpStr = Right("00" & tmpStr, 2) '��ʽ��,���ȱ�Ϊ2 ����ǰ����0��
            S = S & "0" & tmpStr & "H,"
        Next
        
        For I = 17 To 24
            tmpStr = Hex(ZMVH(I2, I))
            tmpStr = Right("00" & tmpStr, 2) '��ʽ��,���ȱ�Ϊ2 ����ǰ����0��
            S = S & "0" & tmpStr & "H,"
        Next
        CharN = CharN + 1
        'S = S & ";" & HzStr1(CharN) & "-" & I2 & vbCrLf & "DB    "
        S = S & ";" & HzStr1(CharN) & "-" & CharN & vbCrLf & "DB    "
        For I = 9 To 16
            tmpStr = Hex(ZMVH(I2, I))
            tmpStr = Right("00" & tmpStr, 2) '��ʽ��,���ȱ�Ϊ2 ����ǰ����0��
            S = S & "0" & tmpStr & "H,"
        Next
        
        For I = 25 To 32
            tmpStr = Hex(ZMVH(I2, I))
            tmpStr = Right("00" & tmpStr, 2) '��ʽ��,���ȱ�Ϊ2 ����ǰ����0��
            S = S & "0" & tmpStr & "H,"
        Next
        
        'CharN = CharN + 1
        'S = S & ";" & HzStr1(CharN) & "-" & I2 & vbCrLf & "DB    "
        
        CharN = CharN + 1
        If (CharN > UBound(HzStr1)) Then
            sTmp = " "
        Else
            sTmp = HzStr1(CharN)
        End If
'        S = S & ";" & sTmp & "-" & I2 & vbCrLf
        S = S & ";" & sTmp & "-" & CharN & vbCrLf
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

Function GetAsc8_16(Char As String) As String


Dim Asc8_16(95) As String
Asc8_16(0) = "&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00"     '&H20' '
Asc8_16(1) = "&H00,&H00,&H00,&H18,&H3C,&H3C,&H3C,&H18,&H18,&H18,&H00,&H18,&H18,&H00,&H00,&H00"     '&H21'!'
Asc8_16(2) = "&H00,&H00,&H66,&H66,&H66,&H24,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00"     '&H22'"'
Asc8_16(3) = "&H00,&H00,&H00,&H00,&H6C,&H6C,&HFE,&H6C,&H6C,&H6C,&HFE,&H6C,&H6C,&H00,&H00,&H00"     '&H23'#'
Asc8_16(4) = "&H00,&H18,&H18,&H7C,&HC6,&HC2,&HC0,&H7C,&H06,&H06,&H86,&HC6,&H7C,&H18,&H18,&H00"     '&H24'$'
Asc8_16(5) = "&H00,&H00,&H00,&H00,&H00,&HC2,&HC6,&H0C,&H18,&H30,&H60,&HC6,&H86,&H00,&H00,&H00"     '&H25'%'
Asc8_16(6) = "&H00,&H00,&H00,&H38,&H6C,&H6C,&H38,&H76,&HDC,&HCC,&HCC,&HCC,&H76,&H00,&H00,&H00"     '&H26'&'
Asc8_16(7) = "&H00,&H00,&H30,&H30,&H30,&H60,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00"     '&H27'''
Asc8_16(8) = "&H00,&H00,&H00,&H0C,&H18,&H30,&H30,&H30,&H30,&H30,&H30,&H18,&H0C,&H00,&H00,&H00"     '&H28'('
Asc8_16(9) = "&H00,&H00,&H00,&H30,&H18,&H0C,&H0C,&H0C,&H0C,&H0C,&H0C,&H18,&H30,&H00,&H00,&H00"     '&H29')'
Asc8_16(10) = "&H00,&H00,&H00,&H00,&H00,&H00,&H66,&H3C,&HFF,&H3C,&H66,&H00,&H00,&H00,&H00,&H00"     '&H2A'*'
Asc8_16(11) = "&H00,&H00,&H00,&H00,&H00,&H00,&H18,&H18,&H7E,&H18,&H18,&H00,&H00,&H00,&H00,&H00"     '&H2B'+'
Asc8_16(12) = "&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H18,&H18,&H18,&H30,&H00,&H00"     '&H2C','
Asc8_16(13) = "&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&HFE,&H00,&H00,&H00,&H00,&H00,&H00,&H00"     '&H2D'-'
Asc8_16(14) = "&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H18,&H18,&H00,&H00,&H00"     '&H2E'.'
Asc8_16(15) = "&H00,&H00,&H00,&H00,&H00,&H02,&H06,&H0C,&H18,&H30,&H60,&HC0,&H80,&H00,&H00,&H00"     '&H2F'/'
Asc8_16(16) = "&H00,&H00,&H00,&H38,&H6C,&HC6,&HC6,&HD6,&HD6,&HC6,&HC6,&H6C,&H38,&H00,&H00,&H00"     '&H30'0'
Asc8_16(17) = "&H00,&H00,&H00,&H18,&H38,&H78,&H18,&H18,&H18,&H18,&H18,&H18,&H7E,&H00,&H00,&H00"     '&H31'1'
Asc8_16(18) = "&H00,&H00,&H00,&H7C,&HC6,&H06,&H0C,&H18,&H30,&H60,&HC0,&HC6,&HFE,&H00,&H00,&H00"     '&H32'2'
Asc8_16(19) = "&H00,&H00,&H00,&H7C,&HC6,&H06,&H06,&H3C,&H06,&H06,&H06,&HC6,&H7C,&H00,&H00,&H00"     '&H33'3'
Asc8_16(20) = "&H00,&H00,&H00,&H0C,&H1C,&H3C,&H6C,&HCC,&HFE,&H0C,&H0C,&H0C,&H1E,&H00,&H00,&H00"     '&H34'4'
Asc8_16(21) = "&H00,&H00,&H00,&HFE,&HC0,&HC0,&HC0,&HFC,&H06,&H06,&H06,&HC6,&H7C,&H00,&H00,&H00"     '&H35'5'
Asc8_16(22) = "&H00,&H00,&H00,&H38,&H60,&HC0,&HC0,&HFC,&HC6,&HC6,&HC6,&HC6,&H7C,&H00,&H00,&H00"     '&H36'6'
Asc8_16(23) = "&H00,&H00,&H00,&HFE,&HC6,&H06,&H06,&H0C,&H18,&H30,&H30,&H30,&H30,&H00,&H00,&H00"     '&H37'7'
Asc8_16(24) = "&H00,&H00,&H00,&H7C,&HC6,&HC6,&HC6,&H7C,&HC6,&HC6,&HC6,&HC6,&H7C,&H00,&H00,&H00"     '&H38'8'
Asc8_16(25) = "&H00,&H00,&H00,&H7C,&HC6,&HC6,&HC6,&H7E,&H06,&H06,&H06,&H0C,&H78,&H00,&H00,&H00"     '&H39'9'
Asc8_16(26) = "&H00,&H00,&H00,&H00,&H00,&H18,&H18,&H00,&H00,&H00,&H18,&H18,&H00,&H00,&H00,&H00"     '&H3A':'
Asc8_16(27) = "&H00,&H00,&H00,&H00,&H00,&H18,&H18,&H00,&H00,&H00,&H18,&H18,&H30,&H00,&H00,&H00"     '&H3B';'
Asc8_16(28) = "&H00,&H00,&H00,&H00,&H06,&H0C,&H18,&H30,&H60,&H30,&H18,&H0C,&H06,&H00,&H00,&H00"     '&H3C'<'
Asc8_16(29) = "&H00,&H00,&H00,&H00,&H00,&H00,&H7E,&H00,&H00,&H7E,&H00,&H00,&H00,&H00,&H00,&H00"     '&H3D'='
Asc8_16(30) = "&H00,&H00,&H00,&H00,&H60,&H30,&H18,&H0C,&H06,&H0C,&H18,&H30,&H60,&H00,&H00,&H00"     '&H3E'>'
Asc8_16(31) = "&H00,&H00,&H00,&H7C,&HC6,&HC6,&H0C,&H18,&H18,&H18,&H00,&H18,&H18,&H00,&H00,&H00"     '&H3F'?'
Asc8_16(32) = "&H00,&H00,&H00,&H00,&H7C,&HC6,&HC6,&HDE,&HDE,&HDE,&HDC,&HC0,&H7C,&H00,&H00,&H00"     '&H40'@'
Asc8_16(33) = "&H00,&H00,&H00,&H10,&H38,&H6C,&HC6,&HC6,&HFE,&HC6,&HC6,&HC6,&HC6,&H00,&H00,&H00"     '&H41'A'
Asc8_16(34) = "&H00,&H00,&H00,&HFC,&H66,&H66,&H66,&H7C,&H66,&H66,&H66,&H66,&HFC,&H00,&H00,&H00"     '&H42'B'
Asc8_16(35) = "&H00,&H00,&H00,&H3C,&H66,&HC2,&HC0,&HC0,&HC0,&HC0,&HC2,&H66,&H3C,&H00,&H00,&H00"     '&H43'C'
Asc8_16(36) = "&H00,&H00,&H00,&HF8,&H6C,&H66,&H66,&H66,&H66,&H66,&H66,&H6C,&HF8,&H00,&H00,&H00"     '&H44'D'
Asc8_16(37) = "&H00,&H00,&H00,&HFE,&H66,&H62,&H68,&H78,&H68,&H60,&H62,&H66,&HFE,&H00,&H00,&H00"     '&H45'E'
Asc8_16(38) = "&H00,&H00,&H00,&HFE,&H66,&H62,&H68,&H78,&H68,&H60,&H60,&H60,&HF0,&H00,&H00,&H00"     '&H46'F'
Asc8_16(39) = "&H00,&H00,&H00,&H3C,&H66,&HC2,&HC0,&HC0,&HDE,&HC6,&HC6,&H66,&H3A,&H00,&H00,&H00"     '&H47'G'
Asc8_16(40) = "&H00,&H00,&H00,&HC6,&HC6,&HC6,&HC6,&HFE,&HC6,&HC6,&HC6,&HC6,&HC6,&H00,&H00,&H00"     '&H48'H'
Asc8_16(41) = "&H00,&H00,&H00,&H3C,&H18,&H18,&H18,&H18,&H18,&H18,&H18,&H18,&H3C,&H00,&H00,&H00"     '&H49'I'
Asc8_16(42) = "&H00,&H00,&H00,&H1E,&H0C,&H0C,&H0C,&H0C,&H0C,&HCC,&HCC,&HCC,&H78,&H00,&H00,&H00"     '&H4A'J'
Asc8_16(43) = "&H00,&H00,&H00,&HE6,&H66,&H66,&H6C,&H78,&H78,&H6C,&H66,&H66,&HE6,&H00,&H00,&H00"     '&H4B'K'
Asc8_16(44) = "&H00,&H00,&H00,&HF0,&H60,&H60,&H60,&H60,&H60,&H60,&H62,&H66,&HFE,&H00,&H00,&H00"     '&H4C'L'
Asc8_16(45) = "&H00,&H00,&H00,&HC6,&HEE,&HFE,&HFE,&HD6,&HC6,&HC6,&HC6,&HC6,&HC6,&H00,&H00,&H00"     '&H4D'M'
Asc8_16(46) = "&H00,&H00,&H00,&HC6,&HE6,&HF6,&HFE,&HDE,&HCE,&HC6,&HC6,&HC6,&HC6,&H00,&H00,&H00"     '&H4E'N'
Asc8_16(47) = "&H00,&H00,&H00,&H7C,&HC6,&HC6,&HC6,&HC6,&HC6,&HC6,&HC6,&HC6,&H7C,&H00,&H00,&H00"     '&H4F'O'
Asc8_16(48) = "&H00,&H00,&H00,&HFC,&H66,&H66,&H66,&H7C,&H60,&H60,&H60,&H60,&HF0,&H00,&H00,&H00"     '&H50'P'
Asc8_16(49) = "&H00,&H00,&H00,&H7C,&HC6,&HC6,&HC6,&HC6,&HC6,&HC6,&HD6,&HDE,&H7C,&H0C,&H0E,&H00"     '&H51'Q'
Asc8_16(50) = "&H00,&H00,&H00,&HFC,&H66,&H66,&H66,&H7C,&H6C,&H66,&H66,&H66,&HE6,&H00,&H00,&H00"     '&H52'R'
Asc8_16(51) = "&H00,&H00,&H00,&H7C,&HC6,&HC6,&H60,&H38,&H0C,&H06,&HC6,&HC6,&H7C,&H00,&H00,&H00"     '&H53'S'
Asc8_16(52) = "&H00,&H00,&H00,&H7E,&H7E,&H5A,&H18,&H18,&H18,&H18,&H18,&H18,&H3C,&H00,&H00,&H00"     '&H54'T'
Asc8_16(53) = "&H00,&H00,&H00,&HC6,&HC6,&HC6,&HC6,&HC6,&HC6,&HC6,&HC6,&HC6,&H7C,&H00,&H00,&H00"     '&H55'U'
Asc8_16(54) = "&H00,&H00,&H00,&HC6,&HC6,&HC6,&HC6,&HC6,&HC6,&HC6,&H6C,&H38,&H10,&H00,&H00,&H00"     '&H56'V'
Asc8_16(55) = "&H00,&H00,&H00,&HC6,&HC6,&HC6,&HC6,&HD6,&HD6,&HD6,&HFE,&HEE,&H6C,&H00,&H00,&H00"     '&H57'W'
Asc8_16(56) = "&H00,&H00,&H00,&HC6,&HC6,&H6C,&H7C,&H38,&H38,&H7C,&H6C,&HC6,&HC6,&H00,&H00,&H00"     '&H58'X'
Asc8_16(57) = "&H00,&H00,&H00,&H66,&H66,&H66,&H66,&H3C,&H18,&H18,&H18,&H18,&H3C,&H00,&H00,&H00"     '&H59'Y'
Asc8_16(58) = "&H00,&H00,&H00,&HFE,&HC6,&H86,&H0C,&H18,&H30,&H60,&HC2,&HC6,&HFE,&H00,&H00,&H00"     '&H5A'Z'
Asc8_16(59) = "&H00,&H00,&H00,&H3C,&H30,&H30,&H30,&H30,&H30,&H30,&H30,&H30,&H3C,&H00,&H00,&H00"     '&H5B'['
Asc8_16(60) = "&H00,&H00,&H00,&H00,&H80,&HC0,&HE0,&H70,&H38,&H1C,&H0E,&H06,&H02,&H00,&H00,&H00"     '&H5C'\'
Asc8_16(61) = "&H00,&H00,&H00,&H3C,&H0C,&H0C,&H0C,&H0C,&H0C,&H0C,&H0C,&H0C,&H3C,&H00,&H00,&H00"     '&H5D']'
Asc8_16(62) = "&H00,&H10,&H38,&H6C,&HC6,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00"     '&H5E'^'
Asc8_16(63) = "&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&HFF,&H00"     '&H5F'_'
Asc8_16(64) = "&H00,&H30,&H30,&H18,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00"     '&H60'`'
Asc8_16(65) = "&H00,&H00,&H00,&H00,&H00,&H00,&H78,&H0C,&H7C,&HCC,&HCC,&HCC,&H76,&H00,&H00,&H00"     '&H61'a'
Asc8_16(66) = "&H00,&H00,&H00,&HE0,&H60,&H60,&H78,&H6C,&H66,&H66,&H66,&H66,&H7C,&H00,&H00,&H00"     '&H62'b'
Asc8_16(67) = "&H00,&H00,&H00,&H00,&H00,&H00,&H7C,&HC6,&HC0,&HC0,&HC0,&HC6,&H7C,&H00,&H00,&H00"     '&H63'c'
Asc8_16(68) = "&H00,&H00,&H00,&H1C,&H0C,&H0C,&H3C,&H6C,&HCC,&HCC,&HCC,&HCC,&H76,&H00,&H00,&H00"     '&H64'd'
Asc8_16(69) = "&H00,&H00,&H00,&H00,&H00,&H00,&H7C,&HC6,&HFE,&HC0,&HC0,&HC6,&H7C,&H00,&H00,&H00"     '&H65'e'
Asc8_16(70) = "&H00,&H00,&H00,&H38,&H6C,&H64,&H60,&HF0,&H60,&H60,&H60,&H60,&HF0,&H00,&H00,&H00"     '&H66'f'
Asc8_16(71) = "&H00,&H00,&H00,&H00,&H00,&H00,&H76,&HCC,&HCC,&HCC,&HCC,&HCC,&H7C,&H0C,&HCC,&H78"     '&H67'g'
Asc8_16(72) = "&H00,&H00,&H00,&HE0,&H60,&H60,&H6C,&H76,&H66,&H66,&H66,&H66,&HE6,&H00,&H00,&H00"     '&H68'h'
Asc8_16(73) = "&H00,&H00,&H00,&H18,&H18,&H00,&H38,&H18,&H18,&H18,&H18,&H18,&H3C,&H00,&H00,&H00"     '&H69'i'
Asc8_16(74) = "&H00,&H00,&H00,&H06,&H06,&H00,&H0E,&H06,&H06,&H06,&H06,&H06,&H06,&H66,&H66,&H3C"     '&H6A'j'
Asc8_16(75) = "&H00,&H00,&H00,&HE0,&H60,&H60,&H66,&H6C,&H78,&H78,&H6C,&H66,&HE6,&H00,&H00,&H00"     '&H6B'k'
Asc8_16(76) = "&H00,&H00,&H00,&H38,&H18,&H18,&H18,&H18,&H18,&H18,&H18,&H18,&H3C,&H00,&H00,&H00"     '&H6C'l'
Asc8_16(77) = "&H00,&H00,&H00,&H00,&H00,&H00,&HEC,&HFE,&HD6,&HD6,&HD6,&HD6,&HC6,&H00,&H00,&H00"     '&H6D'm'
Asc8_16(78) = "&H00,&H00,&H00,&H00,&H00,&H00,&HDC,&H66,&H66,&H66,&H66,&H66,&H66,&H00,&H00,&H00"     '&H6E'n'
Asc8_16(79) = "&H00,&H00,&H00,&H00,&H00,&H00,&H7C,&HC6,&HC6,&HC6,&HC6,&HC6,&H7C,&H00,&H00,&H00"     '&H6F'o'
Asc8_16(80) = "&H00,&H00,&H00,&H00,&H00,&H00,&HDC,&H66,&H66,&H66,&H66,&H66,&H7C,&H60,&H60,&HF0"     '&H70'p'
Asc8_16(81) = "&H00,&H00,&H00,&H00,&H00,&H00,&H76,&HCC,&HCC,&HCC,&HCC,&HCC,&H7C,&H0C,&H0C,&H1E"     '&H71'q'
Asc8_16(82) = "&H00,&H00,&H00,&H00,&H00,&H00,&HDC,&H76,&H66,&H60,&H60,&H60,&HF0,&H00,&H00,&H00"     '&H72'r'
Asc8_16(83) = "&H00,&H00,&H00,&H00,&H00,&H00,&H7C,&HC6,&H60,&H38,&H0C,&HC6,&H7C,&H00,&H00,&H00"     '&H73's'
Asc8_16(84) = "&H00,&H00,&H00,&H10,&H30,&H30,&HFC,&H30,&H30,&H30,&H30,&H36,&H1C,&H00,&H00,&H00"     '&H74't'
Asc8_16(85) = "&H00,&H00,&H00,&H00,&H00,&H00,&HCC,&HCC,&HCC,&HCC,&HCC,&HCC,&H76,&H00,&H00,&H00"     '&H75'u'
Asc8_16(86) = "&H00,&H00,&H00,&H00,&H00,&H00,&H66,&H66,&H66,&H66,&H66,&H3C,&H18,&H00,&H00,&H00"     '&H76'v'
Asc8_16(87) = "&H00,&H00,&H00,&H00,&H00,&H00,&HC6,&HC6,&HD6,&HD6,&HD6,&HFE,&H6C,&H00,&H00,&H00"     '&H77'w'
Asc8_16(88) = "&H00,&H00,&H00,&H00,&H00,&H00,&HC6,&H6C,&H38,&H38,&H38,&H6C,&HC6,&H00,&H00,&H00"     '&H78'x'
Asc8_16(89) = "&H00,&H00,&H00,&H00,&H00,&H00,&HC6,&HC6,&HC6,&HC6,&HC6,&HC6,&H7E,&H06,&H0C,&HF8"     '&H79'y'
Asc8_16(90) = "&H00,&H00,&H00,&H00,&H00,&H00,&HFE,&HCC,&H18,&H30,&H60,&HC6,&HFE,&H00,&H00,&H00"     '&H7A'z'
Asc8_16(91) = "&H00,&H00,&H00,&H0E,&H18,&H18,&H18,&H70,&H18,&H18,&H18,&H18,&H0E,&H00,&H00,&H00"     '&H7B'{'
Asc8_16(92) = "&H00,&H00,&H00,&H18,&H18,&H18,&H18,&H00,&H18,&H18,&H18,&H18,&H18,&H00,&H00,&H00"     '&H7C'|'
Asc8_16(93) = "&H00,&H00,&H00,&H70,&H18,&H18,&H18,&H0E,&H18,&H18,&H18,&H18,&H70,&H00,&H00,&H00"     '&H7D'}'
Asc8_16(94) = "&H00,&H00,&H00,&H76,&HDC,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00,&H00"     '&H7E'~'
Asc8_16(95) = "&H00,&H00,&H00,&H00,&H00,&H10,&H38,&H6C,&HC6,&HC6,&HC6,&HFE,&H00,&H00,&H00,&H00"     '&H7F''

    GetAsc8_16 = Asc8_16(Asc(Char) - 32)
End Function


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
'3. Autoadjust (From AMEI)  display Position
'4. Debug and optimize

'Update 2009 - 07 - 16
'Loveyu
'1. Important Support HZK16.GBK!!!!!
'2. Optimize!
