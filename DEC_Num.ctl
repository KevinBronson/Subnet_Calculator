VERSION 5.00
Begin VB.UserControl DEC_Num 
   ClientHeight    =   615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1440
   ScaleHeight     =   615
   ScaleWidth      =   1440
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   0
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "255"
      Top             =   0
      Width           =   915
   End
End
Attribute VB_Name = "DEC_Num"
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
    If Text1 = "" Then
        Value = 0
    Else
        Value = CByte(Text1)
    End If
End Property

Public Property Let Value(bytInput As Byte)
    Dim blnCancel As Boolean
    bytLastValue = Value
    RaiseEvent Change(blnCancel, bytInput, bytLastValue)
    If Not (blnCancel) Then Text1 = CStr(bytInput)
End Property

'Locked
Public Property Get Locked() As Boolean
    Locked = Text1.Locked
End Property

Public Property Let Locked(blnInput As Boolean)
    Text1.Locked = blnInput
End Property

'Enabled
Public Property Get Enabled() As Boolean
    Enabled = Text1.Enabled
End Property

Public Property Let Enabled(blnInput As Boolean)
    Text1.Enabled = blnInput
End Property


'===================================================================================
'                             INTERNAL EVENTS
'===================================================================================

'Text1_KeyPress
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If Len(Text1) > 0 Then
        If OK_Keys(KeyAscii) Then
            If Text1 = "0" Then Text1 = ""
            bytLastValue = Value
        Else
            KeyAscii = 0 'Cancels Input
        End If
    End If
End Sub

'Text1_KeyUp
Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim blnCancel As Boolean
    If Len(Text1) > 0 Then
        Text1 = CLng(Text1) 'Removes leading zeros
        If Not (CLng(Text1) > -1 And CLng(Text1) < 256) Then Text1 = CStr(bytLastValue)
    Else
        Value = 0
        Text1 = 0
    End If
    RaiseEvent UserChange(blnCancel, CByte(Text1), bytLastValue)
    If blnCancel Then Text1 = CStr(bytLastValue)
End Sub

'Text1_LostFocus
Private Sub Text1_LostFocus()
    If Text1 = "" Then Text1 = 0
End Sub

'Text1_Validate
Private Sub Text1_Validate(Cancel As Boolean)
    RaiseEvent Validate(Cancel)
End Sub

'UserControl_Initialize
Private Sub UserControl_Initialize()
    Value = 0
End Sub


'===================================================================================
'                             PRIVATE FUNCTIONS
'===================================================================================

'OK_Keys
Private Function OK_Keys(KeyAscii As Integer) As Boolean
    OK_Keys = (IsNumeric(Chr(KeyAscii)) Or (KeyAscii = 8)) 'KeyAscii = 8 >> Backspace
End Function

















