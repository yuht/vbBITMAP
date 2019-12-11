VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   4680
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim TempData As Byte
    Dim A, B As Integer
    Open App.Path + "/ASC16.bin" For Binary As #1
    Open App.Path + "/ASC16_BitMap.h" For Output As #2
    
    Print #2, Chr(9) & "Dim ASC16(255) As String"
    
    
        For A = 0 To 255
            Print #2, Chr(9) & "ASC16(" & A & ")=""";
            For B = 0 To 15
                Get #1, , TempData
                Print #2, Right("00" & Hex(TempData), 2);
                If B <> 15 Then
                    Print #2, ",";
                End If
            Next
            Print #2, """" & Chr(9) & "'" + Right("000" + Trim((Str(A))), 3) + "'" + Chr(A) + "'"
        Next
    Close #2
    Close #1
    End
End Sub
