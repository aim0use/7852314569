VERSION 5.00
Begin VB.Form frmProjEul1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multiples of 3 and 5 below 1000"
   ClientHeight    =   300
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   300
   ScaleWidth      =   2910
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "frmProjEul1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
For i = 1 To 999
    If i Mod 3 = "0" Or i Mod 5 = "0" Then
        Text1.Text = Val(Text1.Text) + i
    End If
Next
End Sub
