VERSION 5.00
Begin VB.Form frmDragRace 
   Caption         =   "DragRace"
   ClientHeight    =   6900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7725
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6900
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdComp 
      Caption         =   "&Let the Computer Choose"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   6000
      Width           =   2895
   End
   Begin VB.CommandButton cmdSelf 
      Caption         =   "&I Want to Choose"
      Height          =   615
      Left            =   3360
      TabIndex        =   3
      Top             =   6000
      Width           =   2415
   End
   Begin VB.CommandButton cmdTransportation 
      Caption         =   "&Transportation"
      Height          =   615
      Left            =   3360
      TabIndex        =   2
      Top             =   6000
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton cmdJalopies 
      Caption         =   "&Jalopies/Builders"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   6000
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblDis 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
   End
End
Attribute VB_Name = "frmDragRace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdComp_Click()
    Load frmComp
    Unload frmDragRace
End Sub

Private Sub cmdJalopies_Click()
    frmJalopies.Visible = True
    frmDragRace.Visible = False
End Sub

Private Sub cmdSelf_Click()
    Load frmSelf
    Unload (frmDragRace)
End Sub

Private Sub cmdTransportation_Click()
    frmTransportation.Visible = True
    frmDragRace.Visible = False
End Sub

Private Sub Form_Load()
    'lblDis.Caption = "Hello " & Player1 & ", you will start with $" & Cash1 & ".  You can choose a car from among the available junkers and jalopies for the basic vehicle.  Everything runs 'as is' but may need work soon or you will blow the engine.  Various mechanical parts will be available to you and you may have an opportunity to perform body work that might improve the value of your project."
    'lblDis.Caption = lblDis.Caption & vbCrLf & "     "
    'lblDis.Caption = lblDis.Caption & "Watch for the phrase 'Engine Work Needed' to tell which engines are likely to self-destruct soon.  The computer will give you the cost of body and paint repairs if they are needed to attain 'full value'."
    'lblDis.Caption = lblDis.Caption & vbCrLf & vbCrLf & "Press the button of where you would like to go next."
    lblDis.Caption = "Hello " & PlayerName(Player) & ", you will start with $" & Cash(Player) & ".  You can choose a car from" '; N$(L)
    lblDis.Caption = lblDis.Caption & " the back row of a disreputable used car lot (prices less than $200).  Or you    can let the computer choose a car for you from the front row of a junk yard.  Junk yard prices are less than $50."
End Sub

