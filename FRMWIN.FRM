VERSION 5.00
Begin VB.Form frmWin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Congradulations"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3105
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   3105
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   360
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label labWin 
      Alignment       =   2  'Center
      Caption         =   "You Win"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
End
Attribute VB_Name = "frmWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Unload Me
    
End Sub

Private Sub Timer1_Timer()

    If labWin.ForeColor = QBColor(0) Then
        labWin.ForeColor = QBColor(14)
    Else
        labWin.ForeColor = QBColor(0)
    End If

End Sub
