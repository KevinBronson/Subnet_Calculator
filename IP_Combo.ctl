VERSION 5.00
Begin VB.UserControl IP_Combo 
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4155
   ScaleHeight     =   1500
   ScaleWidth      =   4155
   Begin SubnetCalculator.IP_Decimal IP_Decimal1 
      Height          =   285
      Left            =   90
      TabIndex        =   1
      Top             =   0
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   503
   End
   Begin SubnetCalculator.IP_Binary IP_Binary1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
   End
End
Attribute VB_Name = "IP_Combo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'============================================================================
'           Copyright© 2002 Kevin Bronson. All rights reserved.
'============================================================================

Public WithEvents IP_Address As IP_Address
Attribute IP_Address.VB_VarHelpID = -1
Public Event Change(strErrorMessage As String)


'Locked
Public Property Get Locked() As Boolean
    Locked = IP_Decimal1.Locked
End Property

'Locked
Public Property Let Locked(blnInput As Boolean)
    IP_Decimal1.Locked = blnInput
End Property

'IP_Decimal1_Change
Private Sub IP_Decimal1_Change(strErrorMessage As String)
    RaiseEvent Change(strErrorMessage)
End Sub

'UserControl_Initialize
Private Sub UserControl_Initialize()
    Set IP_Address = New IP_Address
    Set IP_Decimal1.IP_Address = IP_Address
    Set IP_Binary1.IP_Address = IP_Address
End Sub

'UpdateDisplay
Public Sub UpdateDisplay()
    IP_Decimal1.UpdateDisplay
    IP_Binary1.UpdateDisplay
End Sub



