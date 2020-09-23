VERSION 5.00
Begin VB.Form frmAddWord 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Word"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   3855
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraAddWord 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.CommandButton cmdClear 
         Cancel          =   -1  'True
         Caption         =   "&Clear"
         Height          =   375
         Left            =   1320
         TabIndex        =   7
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton cmdAddWord 
         Caption         =   "&Add Word"
         Default         =   -1  'True
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtAddWord 
         Height          =   285
         Left            =   2880
         TabIndex        =   2
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblWord 
         Height          =   255
         Left            =   1200
         TabIndex        =   4
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label lblNewWord 
         Caption         =   "New Word:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblAddWord 
         Caption         =   "Enter word and click &Add Word:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmAddWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
    lblWord = ""
End Sub
Private Sub cmdOK_Click()
    Unload Me
End Sub
Private Sub cmdAddWord_Click()
    Open App.Path & "/hangman.sys" For Append As #1
        Write #1, lblWord.Caption
    Close #1
    lblWord.Caption = ""
End Sub
Private Sub txtAddWord_Change()
    If txtAddWord.Text = "a" Or txtAddWord = "A" Then
        lblWord.Caption = lblWord.Caption + "A"
        txtAddWord.Text = ""
    End If
    If txtAddWord.Text = "b" Or txtAddWord = "B" Then
        lblWord.Caption = lblWord.Caption + "B"
        txtAddWord.Text = ""
    End If
    If txtAddWord.Text = "c" Or txtAddWord = "C" Then
        lblWord.Caption = lblWord.Caption + "C"
        txtAddWord.Text = ""
    End If
    If txtAddWord.Text = "d" Or txtAddWord = "D" Then
        lblWord.Caption = lblWord.Caption + "D"
        txtAddWord.Text = ""
    End If
    If txtAddWord.Text = "e" Or txtAddWord = "E" Then
        lblWord.Caption = lblWord.Caption + "E"
        txtAddWord.Text = ""
    End If
    If txtAddWord.Text = "f" Or txtAddWord = "F" Then
        lblWord.Caption = lblWord.Caption + "F"
        txtAddWord.Text = ""
    End If
    If txtAddWord.Text = "g" Or txtAddWord = "G" Then
        lblWord.Caption = lblWord.Caption + "G"
        txtAddWord.Text = ""
    End If
    If txtAddWord.Text = "h" Or txtAddWord = "H" Then
        lblWord.Caption = lblWord.Caption + "H"
        txtAddWord.Text = ""
    End If
    If txtAddWord.Text = "i" Or txtAddWord = "I" Then
        lblWord.Caption = lblWord.Caption + "I"
        txtAddWord.Text = ""
    End If
    If txtAddWord.Text = "j" Or txtAddWord = "J" Then
        lblWord.Caption = lblWord.Caption + "J"
        txtAddWord.Text = ""
    End If
    If txtAddWord.Text = "k" Or txtAddWord = "K" Then
        lblWord.Caption = lblWord.Caption + "K"
        txtAddWord.Text = ""
    End If
    If txtAddWord.Text = "l" Or txtAddWord = "L" Then
        lblWord.Caption = lblWord.Caption + "L"
        txtAddWord.Text = ""
    End If
    If txtAddWord.Text = "m" Or txtAddWord = "M" Then
        lblWord.Caption = lblWord.Caption + "M"
        txtAddWord.Text = ""
    End If
    If txtAddWord.Text = "n" Or txtAddWord = "N" Then
        lblWord.Caption = lblWord.Caption + "N"
        txtAddWord.Text = ""
    End If
    If txtAddWord.Text = "o" Or txtAddWord = "O" Then
        lblWord.Caption = lblWord.Caption + "O"
        txtAddWord.Text = ""
    End If
    If txtAddWord.Text = "p" Or txtAddWord = "P" Then
        lblWord.Caption = lblWord.Caption + "P"
        txtAddWord.Text = ""
    End If
    If txtAddWord.Text = "q" Or txtAddWord = "Q" Then
        lblWord.Caption = lblWord.Caption + "Q"
        txtAddWord.Text = ""
    End If
    If txtAddWord.Text = "r" Or txtAddWord = "R" Then
        lblWord.Caption = lblWord.Caption + "R"
        txtAddWord.Text = ""
    End If
    If txtAddWord.Text = "s" Or txtAddWord = "S" Then
        lblWord.Caption = lblWord.Caption + "S"
        txtAddWord.Text = ""
    End If
    If txtAddWord.Text = "t" Or txtAddWord = "T" Then
        lblWord.Caption = lblWord.Caption + "T"
        txtAddWord.Text = ""
    End If
    If txtAddWord.Text = "u" Or txtAddWord = "U" Then
        lblWord.Caption = lblWord.Caption + "U"
        txtAddWord.Text = ""
    End If
    If txtAddWord.Text = "v" Or txtAddWord = "V" Then
        lblWord.Caption = lblWord.Caption + "V"
        txtAddWord.Text = ""
    End If
    If txtAddWord.Text = "w" Or txtAddWord = "W" Then
        lblWord.Caption = lblWord.Caption + "W"
        txtAddWord.Text = ""
    End If
    If txtAddWord.Text = "x" Or txtAddWord = "X" Then
        lblWord.Caption = lblWord.Caption + "X"
        txtAddWord.Text = ""
    End If
    If txtAddWord.Text = "y" Or txtAddWord = "Y" Then
        lblWord.Caption = lblWord.Caption + "Y"
        txtAddWord.Text = ""
    End If
    If txtAddWord.Text = "z" Or txtAddWord = "Z" Then
        lblWord.Caption = lblWord.Caption + "Z"
        txtAddWord.Text = ""
    End If
End Sub
