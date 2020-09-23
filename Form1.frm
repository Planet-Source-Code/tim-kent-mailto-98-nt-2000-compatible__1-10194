VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   525
      Left            =   2730
      TabIndex        =   1
      Top             =   1800
      Width           =   1635
   End
   Begin VB.TextBox Text1 
      Height          =   525
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1830
      Width           =   1515
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

        If Text1.Text <> "" Then
            Call ShellExecute(0&, vbNullString, "mailto: " & Text1.Text, vbNullString, vbNullString, vbNormalFocus)
        Else
            MsgBox "Please enter an e-mail address", vbOKOnly + vbExclamation
        End If
        
End Sub
