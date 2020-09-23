VERSION 5.00
Begin VB.Form frmReg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registration"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   3855
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraInfo 
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   3615
      Begin VB.Label lblEmail 
         Alignment       =   2  'Center
         Caption         =   "ciphersoftware@hotmail.com"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label lblWarn 
         Alignment       =   2  'Center
         Caption         =   "Type in the key exactly as you got it."
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         Caption         =   "www.ciphersoftware.cjb.net"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Frame fraReg 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   840
         TabIndex        =   1
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox txtKey 
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Top             =   1680
         Width           =   2415
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   720
         TabIndex        =   3
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lblName 
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblKey 
         Caption         =   "Key:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1680
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    Dim WarePath As String
    Dim RegKeyName As String
    Dim NewKey As String
    Dim RegToUser As String
    Dim RegName As String

    WarePath = "HKEY_LOCAL_MACHINE\Software\CipherSoftware\Hangman"
    RegKeyName = "RegKey"
    NewKey = "CSIOPFHTERRESQ"
    RegToUser = txtName.Text
    RegName = "RegTo"

    If txtKey = "RFfrJ-EqaTr-G31EP-ISdfL-TIh5Q" Then
        CreateKey WarePath
        SetStringValue WarePath, RegKeyName, NewKey
        SetStringValue WarePath, RegName, RegToUser
        MsgBox "Registration was successfull" + Chr(10) + Chr(13) + Chr(10) + Chr(13) + "Regsitered to: " + txtName.Text + Chr(10) + Chr(13) + "Key: " + txtKey.Text, vbInformation + vbOKOnly, "Thank You!"
        frmGame.mnuGame.Enabled = True
        frmGame.mnuRegister.Visible = False
        frmAbout.lblRegto = "Registered To: " + txtName.Text
        frmGame.Caption = "The Hangman Game"
        Unload Me
    Else
        MsgBox "Registration was unsuccessfull" + Chr(10) + Chr(13) + "Please try again!", vbExclamation + vbOKOnly, "Sorry!"
        txtKey.Text = ""
        txtKey.SetFocus
    End If

End Sub
