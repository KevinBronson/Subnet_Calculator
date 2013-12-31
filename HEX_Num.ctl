VERSION 5.00
Begin VB.UserControl HEX_Num 
   ClientHeight    =   690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1380
   ScaleHeight     =   690
   ScaleWidth      =   1380
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   450
      MaxLength       =   1
      TabIndex        =   1
      Text            =   "F"
      Top             =   0
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   0
      MaxLength       =   1
      TabIndex        =   0
      Text            =   "F"
      Top             =   0
      Width           =   465
   End
End
Attribute VB_Name = "HEX_Num"
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
    Dim str1 As String
    Dim str2 As String
    str1 = Text1
    str2 = Text2
    If str1 = "" Then str1 = 0
    If str2 = "" Then str2 = 0
    Value = CByte(Val("&H" & str1 & str2))
End Property

Public Property Let Value(bytInput As Byte)
    Dim blnCancel As Boolean
    bytLastValue = Value
    RaiseEvent Change(blnCancel, bytInput, bytLastValue)
    If Not (blnCancel) Then
        If bytInput < 16 Then
            Text1 = "0"
            Text2 = Hex$(bytInput)
        Else
            Text1 = Mid(Hex$(bytInput), 1, 1)
            Text2 = Mid(Hex$(bytInput), 2, 1)
        End If
    End If
End Property

'Locked
Public Property Get Locked() As Boolean
    Locked = Text1.Locked
End Property

Public Property Let Locked(blnInput As Boolean)
    Text1.Locked = blnInput
    Text2.Locked = blnInput
End Property

'Enabled
Public Property Get Enabled() As Boolean
    Enabled = Text1.Enabled
End Property

Public Property Let Enabled(blnInput As Boolean)
    Text1.Enabled = blnInput
    Text2.Enabled = blnInput
End Property


'===================================================================================
'                             INTERNAL EVENTS
'===================================================================================

'Text1_KeyPress
Private Sub Text1_KeyPress(KeyAscii As Integer)
    bytLastValue = Value
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If OK_Keys(KeyAscii) And Not (Locked) Then
        Text1 = ""
    Else
        KeyAscii = 0 'Cancels Input
    End If
End Sub

'Text1_KeyUp
Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 39 Then
        '-->
        Text2.SetFocus
    ElseIf KeyCode = 37 Then
        '<--
        'Nada
    Else
        Dim blnCancel As Boolean
        If Text1 = "" Then Text1 = 0
        RaiseEvent UserChange(blnCancel, Value, bytLastValue)
        If blnCancel Then Value = bytLastValue
        Text2.SetFocus
    End If
End Sub

'Text1_Validate
Private Sub Text1_Validate(Cancel As Boolean)
    RaiseEvent Validate(Cancel)
End Sub

'Text2_KeyPress
Private Sub Text2_KeyPress(KeyAscii As Integer)
    bytLastValue = Value
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If OK_Keys(KeyAscii) And Not (Locked) Then
        Text2 = ""
    Else
        KeyAscii = 0 'Cancels Input
    End If
End Sub

'Text2_KeyUp
Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 37 Then
        '<--
        Text1.SetFocus
    Else
        Dim blnCancel As Boolean
        If Text2 = "" Then Text2 = 0
        RaiseEvent UserChange(blnCancel, Value, bytLastValue)
        If blnCancel Then Value = bytLastValue
    End If
End Sub

'Text2_Validate
Private Sub Text2_Validate(Cancel As Boolean)
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
    OK_Keys = (IsNumeric(Chr(KeyAscii)) Or _
               (KeyAscii >= Asc("A") And KeyAscii <= Asc("F")))
End Function









