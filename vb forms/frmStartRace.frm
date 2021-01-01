VERSION 5.00
Begin VB.Form frmStartRace 
   Caption         =   "Start Race"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6510
   Icon            =   "frmStartRace.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   6240
   ScaleWidth      =   6510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Done"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label1 
      Height          =   4695
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   5655
   End
End
Attribute VB_Name = "frmStartRace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer

Private Sub Command1_Click()
End
End Sub
'**************!
'displays cars
'**************!
Private Sub Form_Load()
For i = 1 To 2
    Cash(i) = Cash(i) - Car(i).Cost
    Label1.Caption = Label1.Caption & vbCrLf & vbCrLf
    Label1.Caption = Label1.Caption & PlayerName(i) & ", you have a " & Car(i).HP & " HP, " & Car(i).CID & " CID " & Car(i).des & " and $" & Cash(i) & " remaining"
Next i
End Sub
