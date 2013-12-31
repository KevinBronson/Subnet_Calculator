VERSION 5.00
Begin VB.UserControl BIN_Num 
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1440
   ScaleHeight     =   600
   ScaleWidth      =   1440
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   450
      MaxLength       =   4
      TabIndex        =   1
      Text            =   "1111"
      Top             =   0
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   0
      MaxLength       =   4
      TabIndex        =   0
      Text            =   "1111"
      Top             =   0
      Width           =   465
   End
End
Attribute VB_Name = "BIN_Num"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'============================================================================
'           Copyright© 2002 Kevin Bronson. All rights reserved.
'============================================================================

'===================================================================================
'                               DECLARATIONS
'===================================================================================

Public Event Validate(blnCancel As Boolean)
Attribute Validate.VB_Description = "Fired when validation event is fired for either text box."
Public Event Change(blnCancel As Boolean, ByVal bytNewValue As Byte, ByVal bytOldValue As Byte)
Attribute Change.VB_Description = "Fired when ever the value is changed."
Public Event UserChange(blnCancel As Boolean, ByVal bytNewValue As Byte, ByVal bytOldValue As Byte)
Attribute UserChange.VB_Description = "Fired when a user changes the value via the keyboard."
Public Event Error(strErrorMessage As String)
Attribute Error.VB_Description = "Fired when an error occurs."

Dim bytLastValue As Byte
Dim lngLastText1SelStart As Long
Dim lngLastText2SelStart As Long


'===================================================================================
'                             PUBLIC PROPERTIES
'===================================================================================

'Value
Public Property Get Value() As Byte
Attribute Value.VB_Description = "Current decimal value of control. Can be used to change value of control. Notice that this is a byte data type; the value range is 0 to 255."
    Value = CByte(Val("&H" & BinarytoHex(Text1 & Text2)))
End Property

Public Property Let Value(bytInput As Byte)
    Dim blnCancel As Boolean
    bytLastValue = Value
    RaiseEvent Change(blnCancel, bytInput, bytLastValue)
    If Not (blnCancel) Then
        If bytInput < 16 Then
            Text1 = "0000"
            Text2 = HexToBinary(Hex$(bytInput))
        Else
            Text1 = HexToBinary(Mid(Hex$(bytInput), 1, 1))
            Text2 = HexToBinary(Mid(Hex$(bytInput), 2, 1))
        End If
    End If
End Property

'Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Used to lock or unlock the control."
    Locked = Text1.Locked
End Property

Public Property Let Locked(blnInput As Boolean)
    Text1.Locked = blnInput
    Text2.Locked = blnInput
End Property

'Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Used to enable or disable the control."
    Enabled = Text1.Enabled
End Property

Public Property Let Enabled(blnInput As Boolean)
    Text1.Enabled = blnInput
    Text2.Enabled = blnInput
End Property


'===================================================================================
'                             INTERNAL EVENTS
'===================================================================================

'Text1_GotFocus
Private Sub Text1_GotFocus()
    If lngLastText1SelStart = 5 Then
        'Focus has been set by Text2
        lngLastText1SelStart = 3
    Else
        lngLastText1SelStart = 0
    End If
    Text1.SelStart = lngLastText1SelStart
    Text1.SelLength = 1
End Sub

'Text1_KeyPress
Private Sub Text1_KeyPress(KeyAscii As Integer)
    bytLastValue = Value
    If Text1.SelLength <> 1 Then Text1.SelLength = 1
    If OK_Keys(KeyAscii) And Not (Locked) Then
        'Proceed
    Else
        KeyAscii = 0 'Cancels Input
    End If
End Sub

'Text1_KeyUp
Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not (Locked) Then
        If KeyCode = 37 Then
            '<--
            lngLastText1SelStart = lngLastText1SelStart - 1
            If lngLastText1SelStart < 0 Then lngLastText1SelStart = 0
            Text1.SelStart = lngLastText1SelStart
            Text1.SelLength = 1
        ElseIf KeyCode = 39 Then
            '-->
            lngLastText1SelStart = lngLastText1SelStart + 1
            If lngLastText1SelStart > 3 Then
                'Move to Text2
                Text2.SetFocus
            Else
                Text1.SelStart = lngLastText1SelStart
                Text1.SelLength = 1
            End If
        Else
            If OK_Keys(KeyCode) Then lngLastText1SelStart = lngLastText1SelStart + 1
            If lngLastText1SelStart > 3 Then
                'Move to Text2
                lngLastText1SelStart = 3
                Text2.SetFocus
            Else
                Text1.SelStart = lngLastText1SelStart
                Text1.SelLength = 1
            End If
        End If
        Dim blnCancel As Boolean
        RaiseEvent UserChange(blnCancel, Value, bytLastValue)
        If blnCancel Then Value = bytLastValue
    End If
End Sub

'Text1_Validate
Private Sub Text1_Validate(Cancel As Boolean)
    RaiseEvent Validate(Cancel)
End Sub

'Text2_GotFocus
Private Sub Text2_GotFocus()
    lngLastText2SelStart = 0
    Text2.SelStart = lngLastText2SelStart
    Text2.SelLength = 1
End Sub

'Text2_KeyPress
Private Sub Text2_KeyPress(KeyAscii As Integer)
    bytLastValue = Value
    If Text2.SelLength <> 1 Then Text2.SelLength = 1
    If OK_Keys(KeyAscii) And Not (Locked) Then
        'Proceed
    Else
        KeyAscii = 0 'Cancels Input
    End If
End Sub

'Text2_KeyUp
Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not (Locked) Then
        If KeyCode = 37 Then
            '<--
            lngLastText2SelStart = lngLastText2SelStart - 1
            If lngLastText2SelStart > 3 Then lngLastText2SelStart = 3
            If lngLastText2SelStart < 0 Then
                'Move to Text1
                Text1.SetFocus
                lngLastText1SelStart = 5
                Text1.SelStart = lngLastText1SelStart
            Else
                Text2.SelStart = lngLastText2SelStart
                Text2.SelLength = 1
            End If
        ElseIf KeyCode = 39 Then
            '-->
            lngLastText2SelStart = lngLastText2SelStart + 1
            If lngLastText2SelStart > 3 Then lngLastText2SelStart = 3
            Text2.SelStart = lngLastText2SelStart
            Text2.SelLength = 1
        Else
            If OK_Keys(KeyCode) Then lngLastText2SelStart = lngLastText2SelStart + 1
            If lngLastText2SelStart > 3 Then lngLastText2SelStart = 3
            Text2.SelStart = lngLastText2SelStart
            Text2.SelLength = 1
        End If
        Dim blnCancel As Boolean
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
    'KeyCode '96' is number 1 from key pad
    'KeyCode '97' is number 0 from key pad
    OK_Keys = (KeyAscii = Asc("0") Or KeyAscii = Asc("1") Or _
               KeyAscii = 96 Or KeyAscii = 97)
End Function

'BinarytoHex
Private Function BinarytoHex(ByVal strNum As String) As String
    Dim a As Long
    Dim lngNumberOfSets
    Dim arrTemp() As String

    If (Len(strNum) Mod 4) <> 0 Then
        'Needs padding
        Select Case (Len(strNum) Mod 4)
            Case 1
                strNum = "000" & strNum
            Case 2
                strNum = "00" & strNum
            Case 3
                strNum = "0" & strNum
        End Select
    End If
    lngNumberOfSets = Int(Len(strNum) / 4)
    ReDim arrTemp(lngNumberOfSets)
    For a = 1 To lngNumberOfSets
        Select Case Mid(strNum, ((a - 1) * 4) + 1, 4)
            Case "0000"
                arrTemp(a) = "0"
            Case "0001"
                arrTemp(a) = "1"
            Case "0010"
                arrTemp(a) = "2"
            Case "0011"
                arrTemp(a) = "3"
            Case "0100"
                arrTemp(a) = "4"
            Case "0101"
                arrTemp(a) = "5"
            Case "0110"
                arrTemp(a) = "6"
            Case "0111"
                arrTemp(a) = "7"
            Case "1000"
                arrTemp(a) = "8"
            Case "1001"
                arrTemp(a) = "9"
            Case "1010"
                arrTemp(a) = "A"
            Case "1011"
                arrTemp(a) = "B"
            Case "1100"
                arrTemp(a) = "C"
            Case "1101"
                arrTemp(a) = "D"
            Case "1110"
                arrTemp(a) = "E"
            Case "1111"
                arrTemp(a) = "F"
        End Select
    Next
    BinarytoHex = Join(arrTemp)
End Function

'HexToBinary
Private Function HexToBinary(ByVal strHexNum As String) As String
    Dim a As Long
    For a = 1 To Len(strHexNum)
        Select Case Mid(strHexNum, a, 1)
            Case "0"
                HexToBinary = HexToBinary & "0000"
            Case "1"
                HexToBinary = HexToBinary & "0001"
            Case "2"
                HexToBinary = HexToBinary & "0010"
            Case "3"
                HexToBinary = HexToBinary & "0011"
            Case "4"
                HexToBinary = HexToBinary & "0100"
            Case "5"
                HexToBinary = HexToBinary & "0101"
            Case "6"
                HexToBinary = HexToBinary & "0110"
            Case "7"
                HexToBinary = HexToBinary & "0111"
            Case "8"
                HexToBinary = HexToBinary & "1000"
            Case "9"
                HexToBinary = HexToBinary & "1001"
            Case "A"
                HexToBinary = HexToBinary & "1010"
            Case "B"
                HexToBinary = HexToBinary & "1011"
            Case "C"
                HexToBinary = HexToBinary & "1100"
            Case "D"
                HexToBinary = HexToBinary & "1101"
            Case "E"
                HexToBinary = HexToBinary & "1110"
            Case "F"
                HexToBinary = HexToBinary & "1111"
            Case Else
                'Error Has Occurred
                HexToBinary = HexToBinary & "X"
        End Select
    Next
End Function












