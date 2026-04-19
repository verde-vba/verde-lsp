VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   ScaleHeight     =   4170
   ScaleWidth      =   5580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
    MsgBox "Button clicked! Text is: " & Text1.Text
End Sub

Private Sub Form_Load()
    Me.Caption = "My Form"
    Text1.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("Are you sure?", vbYesNo) = vbNo Then
        Cancel = True
    End If
End Sub
