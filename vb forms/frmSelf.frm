VERSION 5.00
Begin VB.Form frmSelf 
   Caption         =   "Choose A Car"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7470
   Icon            =   "frmSelf.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   5010
   ScaleWidth      =   7470
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "View Next Car >>"
      Height          =   495
      Left            =   5640
      TabIndex        =   2
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<< View Previous Car"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Data datJunk0Cat 
      Caption         =   "DragRace.mdb"
      Connect         =   "Access"
      DatabaseName    =   "DragRace.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Tag             =   "DragRace.mdb"
      Top             =   3840
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.CommandButton cmdChoose 
      Caption         =   "Choose This Car"
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblDis 
      Height          =   915
      Left            =   1800
      TabIndex        =   12
      Top             =   2640
      Width           =   4605
   End
   Begin VB.Label lblMoney 
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   2160
      TabIndex        =   11
      Top             =   4320
      Width           =   2775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      DataField       =   "des2"
      DataSource      =   "datJunk0Cat"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   1680
      TabIndex        =   10
      Top             =   480
      Width           =   75
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      DataField       =   "des3"
      DataSource      =   "datJunk0Cat"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   1680
      TabIndex        =   9
      Top             =   840
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Number"
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   0
      TabIndex        =   8
      Top             =   240
      Width           =   705
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      DataField       =   "Num"
      DataSource      =   "datJunk0Cat"
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   840
      TabIndex        =   7
      Top             =   240
      Width           =   60
   End
   Begin VB.Label lblHP 
      DataField       =   "HP"
      DataSource      =   "datJunk0Cat"
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   5640
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label lblCID 
      DataField       =   "CID"
      DataSource      =   "datJunk0Cat"
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   6000
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label lblCost 
      DataField       =   "Cost"
      DataSource      =   "datJunk0Cat"
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   6600
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      DataField       =   "des"
      DataSource      =   "datJunk0Cat"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   1680
      TabIndex        =   3
      Top             =   120
      Width           =   45
   End
End
Attribute VB_Name = "frmSelf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim PicName As String 'name of picture to display
    
        'Dim car As JunkocaT
        'Dim counter As Integer
'**************************************************************************************************
'RecordSource for datJunk0Cat needs to be changed for USED50-A for Player1 and USED50-B for Player2
'**************************************************************************************************

Private Sub cmdChoose_Click()
    ChooseCar
If Player = 1 Then
    Player = 2
    Lane = 2
'frmTransportation.Hide
    frmBegin.Show
Else
    frmStartRace.Show 'Load frmStartRace
End If
Unload Me
End Sub

Private Sub Command1_Click() 'NEXT
Unload frmCar
CarDisplay = 0
'Call Form_Load
'Call Form_Activate
'If lblDis.Caption = "" Then counter = counter - 1
'If counter = 9 Then lblDis = lblDis & "L CHEAP"
datJunk0Cat.Recordset.MoveNext
If (datJunk0Cat.Recordset.EOF) Then datJunk0Cat.Recordset.MoveFirst
ShowCar
End Sub

Private Sub Command2_Click() 'PREVIOUS
Unload frmCar
CarDisplay = 0
'If lblDis.Caption = "" Then counter = counter + 1
'If counter = 1 Then counter = 2
'counter = counter - 2
Rem previous button will go back 2 cars and form load will go
Rem forward 1
'Form_Load

datJunk0Cat.Recordset.MovePrevious
If (datJunk0Cat.Recordset.BOF) Then datJunk0Cat.Recordset.MoveLast
'If counter = 9 Then lblDis = lblDis & "L CHEAP"
ShowCar
'Form_Load
End Sub

Private Sub Form_Activate()
If CarDisplay = 1 Then Exit Sub
'datJunk0Cat.DatabaseName = "DragRace.mdb"
'datJunk0Cat.Refresh
    If Player = 1 Then
        datJunk0Cat.RecordSource = "USED50-A"
        datJunk0Cat.Refresh
    Else
        datJunk0Cat.RecordSource = "USED50-B"
        datJunk0Cat.Refresh
    End If
    ShowCar
End Sub

Sub ShowCar()

    'lblDis.Caption = car.des & " Number " & car.Num & " listed at " & car.HP & " HP, " & car.CID & " CID, will cost you $" & car.Cost & " " & car.des2 & " " & car.des3
    'lblMoney.Caption = Player1 & " -- you have $" & Cash & " left to spend."
    'If counter = 0 Then counter = 1
    
    If CarDisplay = 0 Then
        frmCar.Show
        CarDisplay = 1
    End If
    
    lblDis.Caption = "Listed At:  " & datJunk0Cat.Recordset("HP") & " HP, " & datJunk0Cat.Recordset("CID") & " CID, Will Cost You $" & datJunk0Cat.Recordset("Cost")
    lblMoney.Caption = PlayerName(Player) & " -- you have $" & Cash(Player) & " left to spend."
    'Label1.Caption = Label1.Caption & " " & datJunk0Cat.Recordset("des2") & " " & datJunk0Cat.Recordset("des3")
    
    If Label11.Width > Label1.Width And Label11.Width > Label2.Width Then
        frmSelf.Width = Label11.Left + Label11.Width + 100
    ElseIf Label1.Width > Label2.Width Then
        frmSelf.Width = Label1.Left + Label1.Width + 100
    Else
        frmSelf.Width = Label2.Left + Label2.Width + 100
    End If
    If frmSelf.Width < 7590 Then frmSelf.Width = 7590
    
End Sub
