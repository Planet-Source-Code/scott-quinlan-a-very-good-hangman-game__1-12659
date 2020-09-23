VERSION 5.00
Begin VB.Form frmGame 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Hangman Game"
   ClientHeight    =   5295
   ClientLeft      =   1545
   ClientTop       =   1410
   ClientWidth     =   7095
   Icon            =   "frmGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrRandom 
      Interval        =   10
      Left            =   240
      Top             =   240
   End
   Begin VB.TextBox txtGuess 
      Height          =   285
      Left            =   2280
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   5880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame fraPlay 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   6855
      Begin VB.Timer tmrChooseWord 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   600
         Top             =   240
      End
      Begin VB.Frame fraInfo 
         Caption         =   "Game Information:"
         Height          =   855
         Left            =   2520
         TabIndex        =   33
         Top             =   1440
         Width           =   3975
         Begin VB.Label labWrong 
            Alignment       =   2  'Center
            Caption         =   "0"
            Height          =   255
            Left            =   2400
            TabIndex        =   37
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Wrong clicks"
            Height          =   255
            Left            =   2400
            TabIndex        =   36
            Top             =   240
            Width           =   975
         End
         Begin VB.Label labTotal 
            Alignment       =   2  'Center
            Caption         =   "0"
            Height          =   255
            Left            =   600
            TabIndex        =   35
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Total clicks"
            Height          =   255
            Left            =   600
            TabIndex        =   34
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Click a letter in this box to guess a letter:"
         Height          =   1095
         Left            =   2520
         TabIndex        =   5
         Top             =   3840
         Width           =   3975
         Begin VB.Label P 
            Caption         =   "P"
            Height          =   255
            Left            =   3000
            TabIndex        =   31
            Top             =   240
            Width           =   135
         End
         Begin VB.Label L 
            Caption         =   "L"
            Height          =   255
            Left            =   2880
            TabIndex        =   30
            Top             =   480
            Width           =   135
         End
         Begin VB.Label K 
            Caption         =   "K"
            Height          =   255
            Left            =   2640
            TabIndex        =   29
            Top             =   480
            Width           =   135
         End
         Begin VB.Label O 
            Caption         =   "O"
            Height          =   255
            Left            =   2760
            TabIndex        =   28
            Top             =   240
            Width           =   135
         End
         Begin VB.Label I 
            Caption         =   "I"
            Height          =   255
            Left            =   2520
            TabIndex        =   27
            Top             =   240
            Width           =   135
         End
         Begin VB.Label J 
            Caption         =   "J"
            Height          =   255
            Left            =   2400
            TabIndex        =   26
            Top             =   480
            Width           =   135
         End
         Begin VB.Label H 
            Caption         =   "H"
            Height          =   255
            Left            =   2160
            TabIndex        =   25
            Top             =   480
            Width           =   135
         End
         Begin VB.Label U 
            Caption         =   "U"
            Height          =   255
            Left            =   2280
            TabIndex        =   24
            Top             =   240
            Width           =   135
         End
         Begin VB.Label M 
            Caption         =   "M"
            Height          =   255
            Left            =   2520
            TabIndex        =   23
            Top             =   720
            Width           =   135
         End
         Begin VB.Label G 
            Caption         =   "G"
            Height          =   255
            Left            =   1920
            TabIndex        =   22
            Top             =   480
            Width           =   135
         End
         Begin VB.Label Y 
            Caption         =   "Y"
            Height          =   255
            Left            =   2040
            TabIndex        =   21
            Top             =   240
            Width           =   135
         End
         Begin VB.Label T 
            Caption         =   "T"
            Height          =   255
            Left            =   1800
            TabIndex        =   20
            Top             =   240
            Width           =   135
         End
         Begin VB.Label N 
            Caption         =   "N"
            Height          =   255
            Left            =   2280
            TabIndex        =   19
            Top             =   720
            Width           =   135
         End
         Begin VB.Label R 
            Caption         =   "R"
            Height          =   255
            Left            =   1560
            TabIndex        =   18
            Top             =   240
            Width           =   135
         End
         Begin VB.Label V 
            Caption         =   "V"
            Height          =   255
            Left            =   1800
            TabIndex        =   17
            Top             =   720
            Width           =   135
         End
         Begin VB.Label X 
            Caption         =   "X"
            Height          =   255
            Left            =   1320
            TabIndex        =   16
            Top             =   720
            Width           =   135
         End
         Begin VB.Label S 
            Caption         =   "S"
            Height          =   255
            Left            =   1200
            TabIndex        =   15
            Top             =   480
            Width           =   135
         End
         Begin VB.Label W 
            Caption         =   "W"
            Height          =   255
            Left            =   1080
            TabIndex        =   14
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Z 
            Caption         =   "Z"
            Height          =   255
            Left            =   1080
            TabIndex        =   13
            Top             =   720
            Width           =   135
         End
         Begin VB.Label Q 
            Caption         =   "Q"
            Height          =   255
            Left            =   840
            TabIndex        =   12
            Top             =   240
            Width           =   135
         End
         Begin VB.Label F 
            Caption         =   "F"
            Height          =   255
            Left            =   1680
            TabIndex        =   11
            Top             =   480
            Width           =   135
         End
         Begin VB.Label E 
            Caption         =   "E"
            Height          =   255
            Left            =   1320
            TabIndex        =   10
            Top             =   240
            Width           =   135
         End
         Begin VB.Label D 
            Caption         =   "D"
            Height          =   255
            Left            =   1440
            TabIndex        =   9
            Top             =   480
            Width           =   135
         End
         Begin VB.Label C 
            Caption         =   "C"
            Height          =   255
            Left            =   1560
            TabIndex        =   8
            Top             =   720
            Width           =   135
         End
         Begin VB.Label B 
            Caption         =   "B"
            Height          =   255
            Left            =   2040
            TabIndex        =   7
            Top             =   720
            Width           =   135
         End
         Begin VB.Label A 
            Caption         =   "A"
            Height          =   255
            Left            =   960
            TabIndex        =   6
            Top             =   480
            Width           =   135
         End
      End
      Begin VB.Frame fraword 
         Caption         =   "Hidden word:"
         Height          =   855
         Left            =   2520
         TabIndex        =   3
         Top             =   360
         Width           =   3975
         Begin VB.Label labWord 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.Frame frame 
         Caption         =   "Letters clicked:"
         Height          =   1095
         Left            =   2520
         TabIndex        =   1
         Top             =   2520
         Width           =   3975
         Begin VB.Label labOutput 
            Height          =   735
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.Label lblcword 
         Height          =   255
         Left            =   360
         TabIndex        =   40
         Top             =   4680
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label labRandom 
         Height          =   375
         Left            =   480
         TabIndex        =   39
         Top             =   4320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Line support 
         Visible         =   0   'False
         X1              =   1200
         X2              =   840
         Y1              =   1200
         Y2              =   1560
      End
      Begin VB.Line larm 
         Visible         =   0   'False
         X1              =   1200
         X2              =   1680
         Y1              =   1680
         Y2              =   2160
      End
      Begin VB.Line rarm 
         Visible         =   0   'False
         X1              =   1680
         X2              =   2160
         Y1              =   2160
         Y2              =   1680
      End
      Begin VB.Line rleg 
         Visible         =   0   'False
         X1              =   1680
         X2              =   2040
         Y1              =   2760
         Y2              =   3240
      End
      Begin VB.Line lleg 
         Visible         =   0   'False
         X1              =   1680
         X2              =   1320
         Y1              =   2760
         Y2              =   3240
      End
      Begin VB.Line body 
         Visible         =   0   'False
         X1              =   1680
         X2              =   1680
         Y1              =   2760
         Y2              =   1920
      End
      Begin VB.Shape head 
         Height          =   495
         Left            =   1320
         Shape           =   3  'Circle
         Top             =   1440
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Line rope 
         Visible         =   0   'False
         X1              =   1680
         X2              =   1680
         Y1              =   1440
         Y2              =   1200
      End
      Begin VB.Line pole2 
         Visible         =   0   'False
         X1              =   1680
         X2              =   840
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line pole 
         Visible         =   0   'False
         X1              =   840
         X2              =   840
         Y1              =   3600
         Y2              =   1200
      End
      Begin VB.Line Ground 
         Visible         =   0   'False
         X1              =   480
         X2              =   1920
         Y1              =   3600
         Y2              =   3600
      End
   End
   Begin VB.Label Label1 
      Caption         =   "To play click New Game from the file menu"
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   120
      Width           =   3135
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuNew 
         Caption         =   "&New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuStop 
         Caption         =   "&Stop Game"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuResume 
         Caption         =   "&Resume Game"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddWord 
         Caption         =   "&Add Word"
      End
      Begin VB.Menu mnuChooseWord 
         Caption         =   "&Choose Word"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuRegister 
         Caption         =   "&Register"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Registered As Boolean
Dim MyGuess As String
Dim Guessed As String
Dim Word As String
Dim cWord
Dim wrong As Single
Dim template As String
Dim Answer As String
Dim endword As String
Dim chance, num As Byte
Private Sub cmdExit_Click()
    End
End Sub
Public Sub letters(MyGuess As String)
    Dim numWords As Integer
    Dim Chooseword As Integer
    Dim OldTemplate As String
    Dim C As String
    Dim num As String
    Dim wordlenth As Integer
    Dim II As Single
    Guessed = labOutput.Caption
    OldTemplate = template
    For II = 1 To Len(Word)
        C = Mid(Word, II, 1)
        If C = MyGuess Then
            Mid(template, II, 1) = C
        End If
    Next II
    chance = chance + 1
    labTotal = labTotal + 1
    labWord.Caption = template
    labOutput.Caption = labOutput.Caption + " " + MyGuess
    If OldTemplate = template Then
        wrong = wrong + 1
        labWrong = wrong
        labOutput.Caption = Guessed & " " & MyGuess
        If wrong = 1 Then Ground.Visible = True
        If wrong = 2 Then pole.Visible = True
        If wrong = 3 Then pole2.Visible = True
        If wrong = 4 Then support.Visible = True
        If wrong = 5 Then rope.Visible = True
        If wrong = 6 Then head.Visible = True
        If wrong = 7 Then body.Visible = True
        If wrong = 8 Then rleg.Visible = True
        If wrong = 9 Then lleg.Visible = True
        If wrong = 10 Then rarm.Visible = True
        If wrong = 11 Then larm.Visible = True
    End If
     If wrong = 12 Then
        Load frmDead
        frmDead.Show
        wrong = 0
    End If
End Sub
Private Sub A_Click()
    Dim MyGuess As String
    A.Visible = False
    MyGuess = A
    letters (MyGuess)
End Sub
Private Sub B_Click()
    Dim MyGuess As String
    B.Visible = False
    MyGuess = B
    letters (MyGuess)
End Sub
Private Sub C_Click()
    Dim MyGuess As String
    C.Visible = False
    MyGuess = C
    letters (MyGuess)
End Sub
Private Sub D_Click()
    Dim MyGuess As String
    D.Visible = False
    MyGuess = D
    letters (MyGuess)
End Sub
Private Sub E_Click()
    Dim MyGuess As String
    E.Visible = False
    MyGuess = E
    letters (MyGuess)
End Sub
Private Sub F_Click()
    Dim MyGuess As String
    F.Visible = False
    MyGuess = F
    letters (MyGuess)
End Sub
Private Sub Form_Load()
    Dim OpenTimes As String
    Dim UsedTimes As String
    Dim NewOpenTimes As String
    Dim Path As String
    Dim Start As String
    Dim RegKey As String
    Dim RegTo As String
    Load frmAbout
    frmAbout.Hide
    UsedTimes = "VBHM"
    Start = "10"
    Path = "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion"
    OpenTimes = GetStringValue("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion", "VBHM")
    RegTo = GetStringValue("HKEY_LOCAL_MACHINE\Software\CipherSoftware\Hangman", "RegTo")
    RegKey = GetStringValue("HKEY_LOCAL_MACHINE\Software\CipherSoftware\Hangman", "RegKey")
    If RegKey = "CSIOPFHTERRESQ" Then
        mnuRegister.Visible = False
        frmAbout.lblRegto = "Registered to: " + RegTo
        Exit Sub
    Else
        If OpenTimes > "0" Then
            NewOpenTimes = OpenTimes - 1
            frmGame.Caption = frmGame.Caption & " - !! You have " + NewOpenTimes + " uses left !!"
            SetStringValue Path, UsedTimes, NewOpenTimes
        ElseIf OpenTimes = "0" Then
            MsgBox "You must register to continue to use this software" + Chr(10) + Chr(13) + "www.ciphersoftware.cjb.net", vbExclamation + vbOKOnly, "Information"
            mnuGame.Enabled = False
            frmGame.Caption = frmGame.Caption & " - !! Register Now !!"
            Exit Sub
        Else
            CreateKey Path
            SetStringValue Path, UsedTimes, Start
            frmGame.Caption = frmGame.Caption & " - !! You have 10 uses left !!"
        End If
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
Private Sub G_Click()
    Dim MyGuess As String
    G.Visible = False
    MyGuess = G
    letters (MyGuess)
End Sub
Private Sub H_Click()
    Dim MyGuess As String
    H.Visible = False
    MyGuess = H
    letters (MyGuess)
End Sub
Private Sub I_Click()
    Dim MyGuess As String
    I.Visible = False
    MyGuess = I
    letters (MyGuess)
End Sub
Private Sub J_Click()
    Dim MyGuess As String
    J.Visible = False
    MyGuess = J
    letters (MyGuess)
End Sub
Private Sub K_Click()
    Dim MyGuess As String
    K.Visible = False
    MyGuess = K
    letters (MyGuess)
End Sub
Private Sub L_Click()
    Dim MyGuess As String
    L.Visible = False
    MyGuess = L
    letters (MyGuess)
End Sub
Private Sub labWord_Change()
    If labWord = Word Then
        Load frmWin
        frmWin.Show
    End If
End Sub
Private Sub M_Click()
    Dim MyGuess As String
    M.Visible = False
    MyGuess = M
    letters (MyGuess)
End Sub
Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub
Private Sub mnuAddWord_Click()
    Load frmAddWord
    frmAddWord.Show
End Sub
Private Sub mnuChooseWord_Click()
    Load frmChooseWord
    frmChooseWord.Show
    Dim II As Integer
    cWord = ""
    Word = ""
    template = ""
    lblcword.Caption = ""
    If lblcword.Caption = "" Then
        MsgBox "Invalid Word", vbInformation + vbOKOnly, "Invalid Word"
    Else
    lblcword.Caption = ""
    For II = 1 To Len(cWord)
       template = template & "?"
    Next II
    labTotal.Caption = "0"
    labWrong.Caption = "0"
    wrong = 0
    labWord = template
    labOutput.Caption = ""
    txtGuess.Visible = True
    fraPlay.Visible = True
    mnuResume.Enabled = False
    mnuStop.Enabled = True
    A.Visible = True
    B.Visible = True
    C.Visible = True
    D.Visible = True
    E.Visible = True
    F.Visible = True
    G.Visible = True
    H.Visible = True
    I.Visible = True
    J.Visible = True
    K.Visible = True
    L.Visible = True
    M.Visible = True
    N.Visible = True
    O.Visible = True
    P.Visible = True
    Q.Visible = True
    R.Visible = True
    S.Visible = True
    T.Visible = True
    U.Visible = True
    V.Visible = True
    W.Visible = True
    X.Visible = True
    Y.Visible = True
    Z.Visible = True
    Ground.Visible = False
    pole.Visible = False
    pole2.Visible = False
    support.Visible = False
    rope.Visible = False
    head.Visible = False
    body.Visible = False
    larm.Visible = False
    rarm.Visible = False
    lleg.Visible = False
    rleg.Visible = False
    End If
End Sub
Private Sub mnuExit_Click()
    End
End Sub
Private Sub mnuNew_Click()
    Dim numWords As Integer
    Dim Chooseword As Integer
    Dim II As Single
    Dim Random As Integer
    template = ""
    labWord.Caption = " "
    Open App.Path & "\hangman.sys" For Input As #1
    Do While Not EOF(1)
        numWords = numWords + 1
        Input #1, Word
    Loop
    Close #1
    Chooseword = Int(numWords * labRandom.Caption)
    Open App.Path & "\hangman.sys" For Input As #1
    For II = 1 To Chooseword
        Input #1, Word
        Next
    Close #1
    For II = 1 To Len(Word)
       template = template & "?"
    Next II
    labTotal.Caption = "0"
    labWrong.Caption = "0"
    wrong = 0
    labWord = template
    labOutput.Caption = ""
    txtGuess.Visible = True
    fraPlay.Visible = True
    mnuResume.Enabled = False
    mnuStop.Enabled = True
    A.Visible = True
    B.Visible = True
    C.Visible = True
    D.Visible = True
    E.Visible = True
    F.Visible = True
    G.Visible = True
    H.Visible = True
    I.Visible = True
    J.Visible = True
    K.Visible = True
    L.Visible = True
    M.Visible = True
    N.Visible = True
    O.Visible = True
    P.Visible = True
    Q.Visible = True
    R.Visible = True
    S.Visible = True
    T.Visible = True
    U.Visible = True
    V.Visible = True
    W.Visible = True
    X.Visible = True
    Y.Visible = True
    Z.Visible = True
    Ground.Visible = False
    pole.Visible = False
    pole2.Visible = False
    support.Visible = False
    rope.Visible = False
    head.Visible = False
    body.Visible = False
    larm.Visible = False
    rarm.Visible = False
    lleg.Visible = False
    rleg.Visible = False
End Sub
Private Sub mnuRegister_Click()
    Load frmReg
    frmReg.Show
End Sub
Private Sub mnuResume_Click()
    fraPlay.Visible = True
    mnuResume.Enabled = False
    mnuStop.Enabled = True
End Sub
Private Sub mnuStop_Click()
    fraPlay.Visible = False
    mnuResume.Enabled = True
    mnuStop.Enabled = False
End Sub
Private Sub N_Click()
    Dim MyGuess As String
    N.Visible = False
    MyGuess = N
    letters (MyGuess)
End Sub
Private Sub O_Click()
    Dim MyGuess As String
    O.Visible = False
    MyGuess = O
    letters (MyGuess)
End Sub
Private Sub P_Click()
    Dim MyGuess As String
    P.Visible = False
    MyGuess = P
    letters (MyGuess)
End Sub
Private Sub Q_Click()
    Dim MyGuess As String
    Q.Visible = False
    MyGuess = Q
    letters (MyGuess)
End Sub
Private Sub R_Click()
    Dim MyGuess As String
    R.Visible = False
    MyGuess = R
    letters (MyGuess)
End Sub
Private Sub S_Click()
    Dim MyGuess As String
    S.Visible = False
    MyGuess = S
    letters (MyGuess)
End Sub
Private Sub T_Click()
    Dim MyGuess As String
    T.Visible = False
    MyGuess = T
    letters (MyGuess)
End Sub
Private Sub tmrChooseWord_Timer()
    If lblcword.Caption = "" Then
        MsgBox "Invalid Word", vbInformation + vbOKOnly, "Information"
    Else
    Word = lblcword.Caption
    Dim II As Integer
    cWord = lblcword.Caption
    If lblcword.Caption = "" Then
        MsgBox "Invalid Word", vbInformation + vbOKOnly, "Invalid Word"
    Else
    lblcword.Caption = ""
    For II = 1 To Len(cWord)
       template = template & "?"
    Next II
    labTotal.Caption = "0"
    labWrong.Caption = "0"
    wrong = 0
    labWord = template
    labOutput.Caption = ""
    txtGuess.Visible = True
    fraPlay.Visible = True
    mnuResume.Enabled = False
    mnuStop.Enabled = True
    A.Visible = True
    B.Visible = True
    C.Visible = True
    D.Visible = True
    E.Visible = True
    F.Visible = True
    G.Visible = True
    H.Visible = True
    I.Visible = True
    J.Visible = True
    K.Visible = True
    L.Visible = True
    M.Visible = True
    N.Visible = True
    O.Visible = True
    P.Visible = True
    Q.Visible = True
    R.Visible = True
    S.Visible = True
    T.Visible = True
    U.Visible = True
    V.Visible = True
    W.Visible = True
    X.Visible = True
    Y.Visible = True
    Z.Visible = True
    Ground.Visible = False
    pole.Visible = False
    pole2.Visible = False
    support.Visible = False
    rope.Visible = False
    head.Visible = False
    body.Visible = False
    larm.Visible = False
    rarm.Visible = False
    lleg.Visible = False
    rleg.Visible = False
    End If
    End If
    tmrChooseWord.Enabled = False
End Sub
Private Sub tmrRandom_Timer()
    labRandom.Caption = Rnd
End Sub
Private Sub U_Click()
    Dim MyGuess As String
    U.Visible = False
    MyGuess = U
    letters (MyGuess)
End Sub
Private Sub V_Click()
    Dim MyGuess As String
    V.Visible = False
    MyGuess = V
    letters (MyGuess)
End Sub
Private Sub W_Click()
    Dim MyGuess As String
    W.Visible = False
    MyGuess = W
    letters (MyGuess)
End Sub
Private Sub X_Click()
    Dim MyGuess As String
    X.Visible = False
    MyGuess = X
    letters (MyGuess)
End Sub
Private Sub Y_Click()
    Dim MyGuess As String
    Y.Visible = False
    MyGuess = Y
    letters (MyGuess)
End Sub
Private Sub Z_Click()
    Dim MyGuess As String
    Z.Visible = False
    MyGuess = Z
    letters (MyGuess)
End Sub
