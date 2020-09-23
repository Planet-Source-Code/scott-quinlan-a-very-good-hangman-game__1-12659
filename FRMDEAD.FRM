VERSION 5.00
Begin VB.Form frmDead 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sorry"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3540
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   3540
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   120
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label labLose 
      Alignment       =   2  'Center
      Caption         =   "You Lose"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmDead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    Unload Me
    
End Sub

Private Sub Timer1_Timer()

    If labLose.ForeColor = QBColor(0) Then
        labLose.ForeColor = QBColor(12)
    Else
        labLose.ForeColor = QBColor(0)
    End If

End Sub

Private Sub tmrCrash_Timer()

    Load frmCrash
    frmCrash.Show

End Sub
