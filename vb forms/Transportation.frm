VERSION 5.00
Begin VB.Form frmTransportation 
   Caption         =   "Transportation"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7335
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   5040
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdChoose 
      Caption         =   "Choose This Car"
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Data datJunk0Cat 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\projects\DragRace\DragRace.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Junk0Cat"
      Top             =   3960
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<< Previous"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Continue >>"
      Height          =   495
      Left            =   6000
      TabIndex        =   0
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label11 
      DataField       =   "des"
      DataSource      =   "datJunk0Cat"
      ForeColor       =   &H00000080&
      Height          =   735
      Left            =   1680
      TabIndex        =   11
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label lblCost 
      DataField       =   "Cost"
      DataSource      =   "datJunk0Cat"
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   6720
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblCID 
      DataField       =   "CID"
      DataSource      =   "datJunk0Cat"
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   6120
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label lblHP 
      DataField       =   "HP"
      DataSource      =   "datJunk0Cat"
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   5760
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      DataField       =   "Num"
      DataSource      =   "datJunk0Cat"
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   960
      TabIndex        =   7
      Top             =   360
      Width           =   60
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Number"
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   705
   End
   Begin VB.Label Label2 
      DataField       =   "des3"
      DataSource      =   "datJunk0Cat"
      ForeColor       =   &H00000080&
      Height          =   855
      Left            =   1680
      TabIndex        =   5
      Top             =   3120
      Width           =   3735
   End
   Begin VB.Label Label1 
      DataField       =   "des2"
      DataSource      =   "datJunk0Cat"
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   1800
      TabIndex        =   4
      Top             =   2280
      Width           =   3735
   End
   Begin VB.Label lblMoney 
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   4440
      Width           =   2775
   End
   Begin VB.Label lblDis 
      Height          =   735
      Left            =   2040
      TabIndex        =   1
      Top             =   1200
      Width           =   3495
   End
End
Attribute VB_Name = "frmTransportation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
        'Dim car As JunkocaT
        Dim counter As Integer
'

Private Sub cmdChoose_Click()
On Error GoTo ErrNull
'sets values of user1 car for rest of game
'Car1.yrx = datJunk0Cat.Recordset("yrx")
'Car1.weight = datJunk0Cat.Recordset("weight")
'Car1.vlcx = datJunk0Cat.Recordset("vlcx")
'Car1.tz = datJunk0Cat.Recordset("tz")
'Car1.TireWid = datJunk0Cat.Recordset("TireWid")
'Car1.TireDiam = datJunk0Cat.Recordset("TireDiam")
'Car1.Num = datJunk0Cat.Recordset("Num")
'Car1.mfgr = datJunk0Cat.Recordset("mfgr")
'Car1.L5 = datJunk0Cat.Recordset("L5")
'Car1.L4 = datJunk0Cat.Recordset("L4")
'Car1.L3 = datJunk0Cat.Recordset("L3")
'Car1.L2 = datJunk0Cat.Recordset("L2")
'Car1.L1 = datJunk0Cat.Recordset("L1")
'Car1.HP = datJunk0Cat.Recordset("HP")
'Car1.Gears = datJunk0Cat.Recordset("Gears")
'Car1.des3 = datJunk0Cat.Recordset("des3")
'Car1.des2 = datJunk0Cat.Recordset("des2")
'Car1.des = datJunk0Cat.Recordset("des")
'Car1.Cost = datJunk0Cat.Recordset("Cost")
'Car1.CID = datJunk0Cat.Recordset("CID")
frmTransportation.Hide
frmStartRace.Show
Exit Sub
ErrNull:
    If Err.Number = 94 Then Resume Next
End Sub

Private Sub Command1_Click() 'NEXT
'Call Form_Load
'Call Form_Activate
'If lblDis.Caption = "" Then counter = counter - 1
'If counter = 9 Then lblDis = lblDis & "L CHEAP"
datJunk0Cat.Recordset.MoveNext
If (datJunk0Cat.Recordset.EOF) Then datJunk0Cat.Recordset.MoveFirst
Form_Load
End Sub

Private Sub Command2_Click() 'PREVIOUS
'If lblDis.Caption = "" Then counter = counter + 1
'If counter = 1 Then counter = 2
'counter = counter - 2
Rem previous button will go back 2 cars and form load will go
Rem forward 1
'Form_Load

datJunk0Cat.Recordset.MovePrevious
If (datJunk0Cat.Recordset.BOF) Then datJunk0Cat.Recordset.MoveLast
'If counter = 9 Then lblDis = lblDis & "L CHEAP"
'Form_Activate
Form_Load
End Sub

Private Sub Form_Load()
'On Error Resume Next
'counter = counter + 1
    'Open "c:\program files\devstudio\vb\projects\dragrace\Junk0cat.txt" For Random As #1 Len = Len(car)
    'Open "\DragRace.xls" For Random As #1 Len = Len(car)
     '   Get #1, counter, car '.Num
    'Close
    'ShowCar
    lblDis.Caption = "Listed At:  " & lblHP.Caption & " HP, " & lblCID.Caption & " CID, Will Cost You $" & lblCost.Caption
    lblMoney.Caption = Player & " -- you have $" & Cash(1) & " left to spend."
End Sub

Sub ShowCar()

    'lblDis.Caption = car.des & " Number " & car.Num & " listed at " & car.HP & " HP, " & car.CID & " CID, will cost you $" & car.Cost & " " & car.des2 & " " & car.des3
    'lblMoney.Caption = Player1 & " -- you have $" & Cash & " left to spend."
    'If counter = 0 Then counter = 1
    
End Sub
