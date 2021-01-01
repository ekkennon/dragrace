VERSION 5.00
Begin VB.Form frmCar 
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9630
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6960
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows Default
   Begin VB.Image imgCar 
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmCar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PicName As String
Option Explicit

Private Sub Form_Load()
On Error GoTo ErrFix
    'frmCar.Picture = frmSelf.datJunk0Cat.Recordset("pic")
    Set imgCar.Picture = LoadPicture(frmSelf.datJunk0Cat.Recordset("pic"))
            'MsgBox frmSelf.datJunk0Cat.Recordset("pic")
            'PicName = "c:\projects\dragrace\" & frmSelf.datJunk0Cat.Recordset("pic")
            'MsgBox PicName
            'frmCar.picCar.Picture = "c:\projects\dragrace\47 Fleetmaster Coupe.bmp" 'PicName
            'frmCar.lblCar.Caption = frmSelf.datJunk0Cat.Recordset("PicDes")
            
    frmCar.Height = imgCar.Height + 500
    frmCar.Caption = PlayerName(Player) & "'s " & frmSelf.datJunk0Cat.Recordset("PicDes")
    Exit Sub
ErrFix:
    If Err.Number = 13 Then
        Exit Sub
        frmCar.Hide
    End If
End Sub
