VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About The Hangman Game"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4350
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4350
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.Label lblRegto 
         Caption         =   "Registered to: not registered"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   3615
      End
      Begin VB.Label Label5 
         Caption         =   "Cipher Software"
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "The Hangman Game"
         Height          =   255
         Left            =   1200
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Version 1.0"
         Height          =   255
         Left            =   1200
         TabIndex        =   3
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Made By Scott Quinlan"
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         Top             =   960
         Width           =   1695
      End
      Begin VB.Line Line5 
         X1              =   720
         X2              =   600
         Y1              =   1320
         Y2              =   960
      End
      Begin VB.Line Line4 
         X1              =   600
         X2              =   480
         Y1              =   960
         Y2              =   1320
      End
      Begin VB.Line Line3 
         X1              =   360
         X2              =   600
         Y1              =   480
         Y2              =   720
      End
      Begin VB.Line Line2 
         X1              =   600
         X2              =   840
         Y1              =   720
         Y2              =   480
      End
      Begin VB.Line Line1 
         X1              =   600
         X2              =   600
         Y1              =   960
         Y2              =   600
      End
      Begin VB.Shape Shape1 
         Height          =   255
         Left            =   480
         Shape           =   3  'Circle
         Top             =   360
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    frmAbout.Hide

End Sub
