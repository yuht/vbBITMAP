VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12270
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   12270
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1035
      Left            =   9840
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   3795
      Left            =   420
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1380
      Width           =   8475
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   420
      TabIndex        =   0
      Text            =   "Â¹"
      Top             =   240
      Width           =   8355
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim zw(1 To 128) As Byte
Dim ledstring As String
Dim AA As String



Private Sub Command1_Click()
  ledstring = Trim(Text1.Text)
  
   AA = Hex(Asc(Mid(ledstring, 1, 1)))
   
    'Text2.Text = AA
    
    bb = (&H5E * (CLng("&H" & Mid(AA, 1, 2)) - &HA1) + (CLng("&H" & Mid(AA, 3, 2)) - &HA1)) * &H20
    'Text2.Text = Text2.Text + Hex(bb)
    
    
    For I = 1 To 32
      Open "HZK16" For Binary As #1
      Get #1, bb + I, zw(I)
     ' Get #1, bb, zw(32 * (j - 1) + I)
      Close #1
    Next I
   
 
    For I = 1 To 32 Step 1
      If (Imod8) = 0 Then
      Text2.Text = Text2.Text & " 0"
     End If
      Text2.Text = Text2.Text & Hex(zw(I)) & "H"
    Next I
  
End Sub

Private Sub Command2_Click()
  End
End Sub

'Private Sub Form_Load()
'Text1.Text = ""
'Text2.Text = ""
'End Sub
