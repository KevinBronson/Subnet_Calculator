VERSION 5.00
Begin VB.UserControl DEC_Num32 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6645
   ScaleHeight     =   3600
   ScaleWidth      =   6645
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "DEC_Num32"
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
    If Text1 = "" Then
        Value = 0
    Else
        Value = CLng(Text1)
    End If
End Property

Public Property Let Value(lngInput As Long)
    Dim blnCancel As Boolean
    lngLastValue = Value
    RaiseEvent Change(blnCancel, lngInput, lngLastValue)
    If Not (blnCancel) Then Text1 = CStr(lngInput)
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
            lngLastValue = Value
        Else
            KeyAscii = 0 'Cancels Input
        End If
    End If
End Sub

'Text1_KeyUp
Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim blnCancel As Boolean
    If Len(Text1) > 0 Then
        Text1 = CDbl(Text1) 'Removes leading zeros
        If Not (CDbl(Text1) > (-2147483649#) And CDbl(Text1) < (2147483648#)) Then Text1 = CStr(lngLastValue)
    Else
        Value = 0
        Text1 = 0
    End If
    RaiseEvent UserChange(blnCancel, CLng(Text1), lngLastValue)
    If blnCancel Then Text1 = CStr(lngLastValue)
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



















