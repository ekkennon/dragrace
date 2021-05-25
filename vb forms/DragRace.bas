Attribute VB_Name = "Module1"
Option Explicit
'Option Base 1
Option Private Module
'************************************
'STARTS WITH frmBegin
'************************************

'Public Player1 As String 'name of player1
'Public Player2 As String 'name of player2
Public PlayerName(1 To 2) As String  'Player(1)=player 1's name, Player(2)=player2's name
Public Player As Integer 'current player number - either 1 or 2
Public Lane As Integer 'Lane of current player
Public Phase As Integer 'current phase 1 to 5
'Public PhasePlayerLane(1, 1, 1) 'tells phase, player, lane
'Public SetPPL(5, 2, 2) As Integer     'sets phase, player, lane

Public CarDisplay As Integer '=0 if car has not been displayed yet (frmCar)


'Public Cash1 As Currency 'amount of cash for player1
'Public Cash2 As Currency 'amount of cash for player2
Public Cash(1 To 2) As Currency

'user defined type with a variable for each field of all tables in DragRace.mdb file
Public Type JunkocaT
    Num As Integer
    HP As Integer
    CID As Integer
    Cost As Single
    des As String * 30
    des2 As String * 70
    des3 As String * 75
    weight As Single
    TireDiam As Single
    TireWid As Single
    Gears As Integer
    L4 As Single
    L3 As Single
    L2 As Single
    L1 As Single
    L5 As Single
    mfgr As String * 5
    vlcx As Integer
    yrx As Integer
    tz As String * 2
    pic As String * 30
    PicDes As String * 50
End Type
    'L4=FIRST GEAR IF IT WERE A 4-SPEED
    'L3=FIRST GEAR IF IT WERE A 3-SPEED, SECOND GEAR IF IT WERE A 4-SPEED
    'L2=SECOND GEAR IF IT WERE A 3-SPEED, THIRD GEAR IF IT WERE A 4-SPEED
    'L4=LAST GEAR   L5=REAR END RATIO
    
    
'Public Car1 As JunkocaT 'user1's car
'Public Car2 As JunkocaT 'user2's car
    
Public Car(1 To 2) As JunkocaT


Public Sub ChooseCar()
On Error GoTo ErrNull
    'sets values of user1 car for rest of game
    Car(Player).yrx = frmSelf.datJunk0Cat.Recordset("yrx")
    Car(Player).weight = frmSelf.datJunk0Cat.Recordset("weight")
    Car(Player).vlcx = frmSelf.datJunk0Cat.Recordset("vlcx")
    Car(Player).tz = frmSelf.datJunk0Cat.Recordset("tz")
    Car(Player).TireWid = frmSelf.datJunk0Cat.Recordset("TireWid")
    Car(Player).TireDiam = frmSelf.datJunk0Cat.Recordset("TireDiam")
    Car(Player).Num = frmSelf.datJunk0Cat.Recordset("Num")
    Car(Player).mfgr = frmSelf.datJunk0Cat.Recordset("mfgr")
    Car(Player).L5 = frmSelf.datJunk0Cat.Recordset("L5")
    Car(Player).L4 = frmSelf.datJunk0Cat.Recordset("L4")
    Car(Player).L3 = frmSelf.datJunk0Cat.Recordset("L3")
    Car(Player).L2 = frmSelf.datJunk0Cat.Recordset("L2")
    Car(Player).L1 = frmSelf.datJunk0Cat.Recordset("L1")
    Car(Player).HP = frmSelf.datJunk0Cat.Recordset("HP")
    Car(Player).Gears = frmSelf.datJunk0Cat.Recordset("Gears")
    Car(Player).des3 = frmSelf.datJunk0Cat.Recordset("des3")
    Car(Player).des2 = frmSelf.datJunk0Cat.Recordset("des2")
    Car(Player).des = frmSelf.datJunk0Cat.Recordset("des")
    Car(Player).Cost = frmSelf.datJunk0Cat.Recordset("Cost")
    Car(Player).CID = frmSelf.datJunk0Cat.Recordset("CID")
    Car(Player).PicDes = frmSelf.datJunk0Cat.Recordset("PicDes")
    Car(Player).pic = frmSelf.datJunk0Cat.Recordset("pic")
    Exit Sub
ErrNull:
    If Err.Number = 94 Then Resume Next 'this works
                   '/\ is when 'tz' is left blank
End Sub
