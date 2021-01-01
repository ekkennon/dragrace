VERSION 5.00
Begin VB.Form frmBegin 
   Caption         =   "Welcome"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6300
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Begin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3300
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMe 
      Caption         =   "Let me Choose"
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   2400
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdComp 
      Caption         =   "Let the Computer Choose"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton cmdBegin 
      Caption         =   "Begin"
      Default         =   -1  'True
      Height          =   615
      Left            =   3480
      TabIndex        =   1
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtName 
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label lblDis 
      Caption         =   "Would you rather have the Computer choose a car for you or would you like to choose one for yourself?"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   720
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Label lblName 
      Height          =   1095
      Left            =   840
      TabIndex        =   2
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "frmBegin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'cmdComp still does nothing

Option Explicit
Dim flag As Integer 'tells if first letter has been changed to Uppercase yet
Dim Letter As String 'first letter of user's name
Dim rest As String 'player's name except for first letter
Dim ans As Integer 'used in Sub txtName_KeyPress
Dim i As Integer, j As Integer

Private Sub cmdBegin_Click()
'frmSelf.datJunk0Cat.DatabaseName = "DragRace.mdb"

If txtName.Text = "" Then
    MsgBox "Type your name in the box and push 'Begin'."
    txtName.SetFocus
    Exit Sub
End If
    PlayerName(Player) = txtName.Text
    lblName.Visible = False
    txtName.Visible = False
    cmdBegin.Visible = False
    cmdComp.Visible = True
    cmdMe.Visible = True
    lblDis.Visible = True
End Sub

Private Sub cmdComp_Click()
If Lane = 1 Then
    Randomize Timer
    i = Int(Rnd * 10)
    For j = 0 To i 'may not work.
        frmSelf.datJunk0Cat.RecordSource = "Junk0Cat"
        frmSelf.datJunk0Cat.Refresh
        frmSelf.datJunk0Cat.Recordset.MoveNext
    Next j
    ChooseCar
    Unload Me
    If Player = 1 Then
        Player = 2
        Lane = 2
        'frmTransportation.Hide
        frmBegin.Show
    Else
        frmStartRace.Show 'Load frmStartRace
    End If
    Rem open c:\pf\ds\vb\p\dr\used50-a.txt
    '14400 IF L=1 THEN OPEN "I",#1,"JUNK2.DG1"
ElseIf Lane = 2 Then
    Rem open c:\pf\ds\vb\p\dr\used50-b.txt
    '14410 IF L=2 THEN OPEN "I",#1,"JUNK1.DG1"
    Randomize Timer
    i = Int(Rnd * 10)
    For j = 0 To i 'may not work.
        frmSelf.datJunk0Cat.RecordSource = "Junk0Cat"
        frmSelf.datJunk0Cat.Refresh
        'frmSelf.datJunk0Cat.Recordset ("Num")
    Next j
    ChooseCar
    Unload Me
    If Lane = 1 Then
        Lane = 2
        'frmTransportation.Hide
        frmBegin.Show
    Else
        frmStartRace.Show 'Load frmStartRace
        
    End If
End If
End Sub

Private Sub cmdMe_Click()
    'frmBegin.Visible = False
    'frmDragRace.Visible = True
    '14210 IF L=1 THEN OPEN "I",#1,"USED50-A.DG1"
    '14215 IF L=2 THEN OPEN "I",#1,"USED50-B.DG1"
    frmSelf.Show
    Unload Me
End Sub

Private Sub Form_Load()
    'Lane = 1
    'Cash1 = 250
    flag = 0
    CarDisplay = 0
    'Letter = ""
    'rest = ""
    'ans = 0
    If PlayerName(1) = "" Then StartGame 'if PlayerName(1) has not been set this is the beginning of the game
    'lblName.Caption = "Player in lane " & Lane & " what is your name?"
    Dim i As Integer
    For i = 1 To 2
        lblName.Caption = "Player in lane " & Lane & " what is your name?"
    Next i
End Sub

Private Sub txtName_Change()
    If Right(txtName, 1) = " " Then 'if user tries to enter a second name
        txtName.Enabled = False
        ans = MsgBox("This program does not accept multiple names right now.  Would you like to change the name you've entered?", vbYesNo)
        If ans = vbYes Then
            flag = 0
            txtName.Text = ""
            txtName.Enabled = True
        End If
        Exit Sub
    End If
End Sub

'************************************************!
'changes first letter of user's name to UpperCase
'************************************************!
Private Sub txtName_KeyPress(KeyAscii As Integer)
    
    If flag = 0 Then 'if program hasn't changed first letter yet
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        flag = 1
    Else 'if first letter has been changed
        KeyAscii = Asc(LCase(Chr(KeyAscii)))
    End If
End Sub

Sub StartGame()
    Player = 1
    Phase = 1
    Lane = 1
    For i = 1 To 2
        Cash(i) = 250 'both players start with $250
    Next i
End Sub

