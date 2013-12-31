VERSION 5.00
Begin VB.UserControl LONG_Num 
   ClientHeight    =   3885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8100
   ScaleHeight     =   3885
   ScaleWidth      =   8100
   Begin SubnetCalculator.BYTE_Num BYTE_Num1 
      Height          =   825
      Left            =   0
      TabIndex        =   0
      Top             =   270
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1455
   End
   Begin SubnetCalculator.BYTE_Num BYTE_Num2 
      Height          =   825
      Left            =   900
      TabIndex        =   1
      Top             =   270
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1455
   End
   Begin SubnetCalculator.BYTE_Num BYTE_Num3 
      Height          =   825
      Left            =   1800
      TabIndex        =   2
      Top             =   270
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1455
   End
   Begin SubnetCalculator.BYTE_Num BYTE_Num4 
      Height          =   825
      Left            =   2700
      TabIndex        =   3
      Top             =   270
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1455
   End
   Begin SubnetCalculator.DEC_Num32 DEC_Num32 
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   503
   End
End
Attribute VB_Name = "LONG_Num"
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
Public Event Change(blnCancel As Boolean, ByVal lngNewValue As Long, ByVal lngOldValue As Long)
Public Event UserChange(blnCancel As Boolean, ByVal lngNewValue As Long, ByVal lngOldValue As Long)
Public Event Error(strErrorMessage As String)

Dim lngLastValue As Long


'===================================================================================
'                             PUBLIC PROPERTIES
'===================================================================================

'Value
Public Property Get Value() As Long
    Value = DEC_Num32.Value
End Property

Public Property Let Value(lngInput As Long)
    lngLastValue = Value
    DEC_Num32.Value = lngInput
    Call LongToBytes
End Property

'BinaryValue
Public Property Get BinaryValue() As String
    BinaryValue = LongToBinary(Value)
End Property

Public Property Let BinaryValue(strInput As String)
    If Len(strInput) > 0 Then
        If IsBinaryNumber(strInput) Then
            Value = BinaryToLong(strInput)
        End If
    Else
        BinaryValue = 0
    End If
End Property

'UseAllTabStops
Public Property Get UseAllTabStops() As Boolean
    UseAllTabStops = DEC_Num32.TabStop
End Property

Public Property Let UseAllTabStops(blnInput As Boolean)
    'DEC_Num32.TabStop = blnInput
    BYTE_Num1.TabStop = blnInput
    BYTE_Num2.TabStop = blnInput
    BYTE_Num3.TabStop = blnInput
    BYTE_Num4.TabStop = blnInput
End Property

'Locked
Public Property Get Locked() As Boolean
    Locked = DEC_Num32.Locked
End Property

Public Property Let Locked(blnInput As Boolean)
    DEC_Num32.Locked = blnInput
    BYTE_Num1.Locked = blnInput
    BYTE_Num2.Locked = blnInput
    BYTE_Num3.Locked = blnInput
    BYTE_Num4.Locked = blnInput
End Property

'Enabled
Public Property Get Enabled() As Boolean
    Enabled = DEC_Num32.Enabled
End Property

Public Property Let Enabled(blnInput As Boolean)
    DEC_Num32.Enabled = blnInput
    BYTE_Num1.Enabled = blnInput
    BYTE_Num2.Enabled = blnInput
    BYTE_Num3.Enabled = blnInput
    BYTE_Num4.Enabled = blnInput
End Property


'===================================================================================
'                             INTERNAL EVENTS
'===================================================================================

'UserControl_Initialize
Private Sub UserControl_Initialize()
    
End Sub

'DEC_Num32_UserChange
Private Sub DEC_Num32_UserChange(blnCancel As Boolean, ByVal lngNewValue As Long, ByVal lngOldValue As Long)
    Call LongToBytes
End Sub

'BYTE_Num1_UserChange
Private Sub BYTE_Num1_UserChange(blnCancel As Boolean, ByVal bytNewValue As Byte, ByVal bytOldValue As Byte)
    Call BytesToLong
End Sub

'BYTE_Num2_UserChange
Private Sub BYTE_Num2_UserChange(blnCancel As Boolean, ByVal bytNewValue As Byte, ByVal bytOldValue As Byte)
    Call BytesToLong
End Sub

'BYTE_Num3_UserChange
Private Sub BYTE_Num3_UserChange(blnCancel As Boolean, ByVal bytNewValue As Byte, ByVal bytOldValue As Byte)
    Call BytesToLong
End Sub

'BYTE_Num4_UserChange
Private Sub BYTE_Num4_UserChange(blnCancel As Boolean, ByVal bytNewValue As Byte, ByVal bytOldValue As Byte)
    Call BytesToLong
End Sub



'===================================================================================
'                            PRIVATE FUNCTIONS
'===================================================================================

'LongToBytes
Private Sub LongToBytes()
    Dim strBinaryNumber As String
    strBinaryNumber = LongToBinary(DEC_Num32.Value)
    'Break into four parts
    BYTE_Num1.Value = BinaryToDecimal(Mid(strBinaryNumber, 1, 8))
    BYTE_Num2.Value = BinaryToDecimal(Mid(strBinaryNumber, 9, 8))
    BYTE_Num3.Value = BinaryToDecimal(Mid(strBinaryNumber, 17, 8))
    BYTE_Num4.Value = BinaryToDecimal(Mid(strBinaryNumber, 25, 8))
End Sub

'BytesToLong
Private Sub BytesToLong()
    Dim strBinaryNumber As String
    strBinaryNumber = ByteToBinary(BYTE_Num1.Value) & _
                      ByteToBinary(BYTE_Num2.Value) & _
                      ByteToBinary(BYTE_Num3.Value) & _
                      ByteToBinary(BYTE_Num4.Value)
    
    DEC_Num32.Value = BinaryToLong(strBinaryNumber)
End Sub

'LongToBinary
Private Function LongToBinary(lngInput As Long) As String
    LongToBinary = DecimalToBinary(lngInput)
    LongToBinary = Zeros(32 - Len(LongToBinary)) & LongToBinary
End Function

'BinaryToLong
Private Function BinaryToLong(strInput As String) As Long
    BinaryToLong = BinaryToDecimal(strInput)
End Function

'Zeros
Private Function Zeros(lngNumber As Long) As String
    Dim a As Long
    For a = 1 To lngNumber
        Zeros = Zeros & "0"
    Next
End Function















