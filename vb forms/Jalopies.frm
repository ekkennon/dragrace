VERSION 5.00
Begin VB.Form frmJalopies 
   Caption         =   "Jalopies"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7380
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   7380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "<< Previous"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdCont 
      Caption         =   "Continue >>"
      Default         =   -1  'True
      Height          =   495
      Left            =   6000
      TabIndex        =   0
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label lblMoney 
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   4080
      Width           =   2775
   End
   Begin VB.Label lblDis 
      Height          =   3615
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "frmJalopies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
        'Dim car As JunkocaT
        Dim counter As Integer

Private Sub cmdCont_Click()
Call Form_Load
Call Form_Activate
If lblDis.Caption = "" Then counter = counter - 1
End Sub

Private Sub Command1_Click()
If lblDis.Caption = "" Then counter = counter + 1
If counter = 1 Then counter = 2
counter = counter - 2
Rem previous button will go back 2 cars and form load will go
Rem forward 1
Form_Load
Form_Activate
End Sub

Private Sub Form_Load()
counter = counter + 1
    'Open "c:\program files\devstudio\vb\projects\dragrace\Junk0car.txt" For Random As #1 Len = Len(car)
    'Open "\Junk0car.txt" For Random As #1 Len = Len(car)
        'Get #1, counter, car '.Num
    Close
End Sub

Private Sub Form_Activate()
    'lblDis.Caption = car.des & " Number " & car.Num & " listed at " & car.HP & " HP, " & car.CID & " CID, will cost you $" & car.Cost & " " & car.des2 & " " & car.des3
    'lblMoney.Caption = Player1 & " -- you have $" & Cash & " left to spend."
    If counter = 0 Then counter = 1
End Sub


