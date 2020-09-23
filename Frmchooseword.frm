VERSION 5.00
Begin VB.Form frmChooseWord 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choose Word"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   3855
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraAddWord 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.TextBox txtChooseWord 
         Height          =   285
         Left            =   2880
         TabIndex        =   3
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdChooseWord 
         Caption         =   "&Choose Word"
         Default         =   -1  'True
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblChooseWord 
         Caption         =   "Enter word and click Choose Word:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label lblMakeWord 
         Caption         =   "Word:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblWord 
         Height          =   255
         Left            =   720
         TabIndex        =   4
         Top             =   720
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmChooseWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdChooseWord_Click()
    Dim Word
    frmGame.Enabled = True
    frmGame.lblcword = lblWord.Caption
    frmGame.tmrChooseWord.Enabled = True
    Unload Me
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    frmGame.Enabled = False
End Sub
Private Sub txtChooseWord_Change()
    If txtChooseWord.Text = "a" Or txtChooseWord = "A" Then
        lblWord.Caption = lblWord.Caption + "A"
        txtChooseWord.Text = ""
    End If
    If txtChooseWord.Text = "b" Or txtChooseWord = "B" Then
        lblWord.Caption = lblWord.Caption + "B"
        txtChooseWord.Text = ""
    End If
    If txtChooseWord.Text = "c" Or txtChooseWord = "C" Then
        lblWord.Caption = lblWord.Caption + "C"
        txtChooseWord.Text = ""
    End If
    If txtChooseWord.Text = "d" Or txtChooseWord = "D" Then
        lblWord.Caption = lblWord.Caption + "D"
        txtChooseWord.Text = ""
    End If
    If txtChooseWord.Text = "e" Or txtChooseWord = "E" Then
        lblWord.Caption = lblWord.Caption + "E"
        txtChooseWord.Text = ""
    End If
    If txtChooseWord.Text = "f" Or txtChooseWord = "F" Then
        lblWord.Caption = lblWord.Caption + "F"
        txtChooseWord.Text = ""
    End If
    If txtChooseWord.Text = "g" Or txtChooseWord = "G" Then
        lblWord.Caption = lblWord.Caption + "G"
        txtChooseWord.Text = ""
    End If
    If txtChooseWord.Text = "h" Or txtChooseWord = "H" Then
        lblWord.Caption = lblWord.Caption + "H"
        txtChooseWord.Text = ""
    End If
    If txtChooseWord.Text = "i" Or txtChooseWord = "I" Then
        lblWord.Caption = lblWord.Caption + "I"
        txtChooseWord.Text = ""
    End If
    If txtChooseWord.Text = "j" Or txtChooseWord = "J" Then
        lblWord.Caption = lblWord.Caption + "J"
        txtChooseWord.Text = ""
    End If
    If txtChooseWord.Text = "k" Or txtChooseWord = "K" Then
        lblWord.Caption = lblWord.Caption + "K"
        txtChooseWord.Text = ""
    End If
    If txtChooseWord.Text = "l" Or txtChooseWord = "L" Then
        lblWord.Caption = lblWord.Caption + "L"
        txtChooseWord.Text = ""
    End If
    If txtChooseWord.Text = "m" Or txtChooseWord = "M" Then
        lblWord.Caption = lblWord.Caption + "M"
        txtChooseWord.Text = ""
    End If
    If txtChooseWord.Text = "n" Or txtChooseWord = "N" Then
        lblWord.Caption = lblWord.Caption + "N"
        txtChooseWord.Text = ""
    End If
    If txtChooseWord.Text = "o" Or txtChooseWord = "O" Then
        lblWord.Caption = lblWord.Caption + "O"
        txtChooseWord.Text = ""
    End If
    If txtChooseWord.Text = "p" Or txtChooseWord = "P" Then
        lblWord.Caption = lblWord.Caption + "P"
        txtChooseWord.Text = ""
    End If
    If txtChooseWord.Text = "q" Or txtChooseWord = "Q" Then
        lblWord.Caption = lblWord.Caption + "Q"
        txtChooseWord.Text = ""
    End If
    If txtChooseWord.Text = "r" Or txtChooseWord = "R" Then
        lblWord.Caption = lblWord.Caption + "R"
        txtChooseWord.Text = ""
    End If
    If txtChooseWord.Text = "s" Or txtChooseWord = "S" Then
        lblWord.Caption = lblWord.Caption + "S"
        txtChooseWord.Text = ""
    End If
    If txtChooseWord.Text = "t" Or txtChooseWord = "T" Then
        lblWord.Caption = lblWord.Caption + "T"
        txtChooseWord.Text = ""
    End If
    If txtChooseWord.Text = "u" Or txtChooseWord = "U" Then
        lblWord.Caption = lblWord.Caption + "U"
        txtChooseWord.Text = ""
    End If
    If txtChooseWord.Text = "v" Or txtChooseWord = "V" Then
        lblWord.Caption = lblWord.Caption + "V"
        txtChooseWord.Text = ""
    End If
    If txtChooseWord.Text = "w" Or txtChooseWord = "W" Then
        lblWord.Caption = lblWord.Caption + "W"
        txtChooseWord.Text = ""
    End If
    If txtChooseWord.Text = "x" Or txtChooseWord = "X" Then
        lblWord.Caption = lblWord.Caption + "X"
        txtChooseWord.Text = ""
    End If
    If txtChooseWord.Text = "y" Or txtChooseWord = "Y" Then
        lblWord.Caption = lblWord.Caption + "Y"
        txtChooseWord.Text = ""
    End If
    If txtChooseWord.Text = "z" Or txtChooseWord = "Z" Then
        lblWord.Caption = lblWord.Caption + "Z"
        txtChooseWord.Text = ""
    End If
End Sub
