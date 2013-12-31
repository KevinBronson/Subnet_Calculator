VERSION 5.00
Begin VB.UserControl BYTE_Num 
   ClientHeight    =   1125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1515
   ScaleHeight     =   1125
   ScaleWidth      =   1515
   Begin SubnetCalculator.HEX_Num HEX_Num1 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   270
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   503
   End
   Begin SubnetCalculator.DEC_Num DEC_Num1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   503
   End
   Begin SubnetCalculator.BIN_Num BIN_Num1 
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   540
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   503
   End
End
Attribute VB_Name = "BYTE_Num"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'============================================================================
'           Copyright© 2002 Kevin Bronson. All rights reserved.
'============================================================================

'===================================================================================
'                             DECLARATIONS
'===================================================================================

Public Event Validate(blnCancel As Boolean)
Public Event Change(blnCancel As Boolean, ByVal bytNewValue As Byte, ByVal bytOldValue As Byte)
Public Event UserChange(blnCancel As Boolean, ByVal bytNewValue As Byte, ByVal bytOldValue As Byte)
Public Event Error(strErrorMessage As String)

Dim bytLastValue As Byte


'===================================================================================
'                             PUBLIC PROPERTIES
'===================================================================================

'Value
Public Property Get Value() As Byte
    Value = DEC_Num1.Value
End Property

Public Property Let Value(bytInput As Byte)
    bytLastValue = Value
    DEC_Num1.Value = bytInput
    HEX_Num1.Value = bytInput
    BIN_Num1.Value = bytInput
End Property

'UseAllTabStops
Public Property Get UseAllTabStops() As Boolean
    UseAllTabStops = HEX_Num1.TabStop
End Property

Public Property Let UseAllTabStops(blnInput As Boolean)
    HEX_Num1.TabStop = blnInput
    BIN_Num1.TabStop = blnInput
End Property

'Locked
Public Property Get Locked() As Boolean
    Locked = DEC_Num1.Locked
End Property

Public Property Let Locked(blnInput As Boolean)
    DEC_Num1.Locked = blnInput
    HEX_Num1.Locked = blnInput
    BIN_Num1.Locked = blnInput
End Property

'Enabled
Public Property Get Enabled() As Boolean
    Enabled = DEC_Num1.Enabled
End Property

Public Property Let Enabled(blnInput As Boolean)
    DEC_Num1.Enabled = blnInput
    HEX_Num1.Enabled = blnInput
    BIN_Num1.Enabled = blnInput
End Property


'===================================================================================
'                             INTERNAL EVENTS
'===================================================================================

'DEC_Num1_UserChange
Private Sub DEC_Num1_UserChange(blnCancel As Boolean, ByVal bytNewValue As Byte, ByVal bytOldValue As Byte)
    HEX_Num1.Value = DEC_Num1.Value
    BIN_Num1.Value = DEC_Num1.Value
    RaiseEvent UserChange(blnCancel, bytNewValue, bytOldValue)
End Sub

'HEX_Num1_UserChange
Private Sub HEX_Num1_UserChange(blnCancel As Boolean, ByVal bytNewValue As Byte, ByVal bytOldValue As Byte)
    DEC_Num1.Value = HEX_Num1.Value
    BIN_Num1.Value = HEX_Num1.Value
    RaiseEvent UserChange(blnCancel, bytNewValue, bytOldValue)
End Sub

'BIN_Num1_UserChange
Private Sub BIN_Num1_UserChange(blnCancel As Boolean, ByVal bytNewValue As Byte, ByVal bytOldValue As Byte)
    DEC_Num1.Value = BIN_Num1.Value
    HEX_Num1.Value = BIN_Num1.Value
    RaiseEvent UserChange(blnCancel, bytNewValue, bytOldValue)
End Sub

'UserControl_Initialize
Private Sub UserControl_Initialize()
    Value = 0
End Sub






