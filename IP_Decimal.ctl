VERSION 5.00
Begin VB.UserControl IP_Decimal 
   ClientHeight    =   1545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3870
   ScaleHeight     =   1545
   ScaleWidth      =   3870
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1350
      TabIndex        =   6
      Text            =   "0"
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   900
      TabIndex        =   4
      Text            =   "255"
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   450
      TabIndex        =   2
      Text            =   "255"
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "255"
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1260
      TabIndex        =   5
      Top             =   0
      Width           =   105
   End
   Begin VB.Label Label1 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   810
      TabIndex        =   3
      Top             =   0
      Width           =   105
   End
   Begin VB.Label Label2 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   0
      Width           =   105
   End
End
Attribute VB_Name = "IP_Decimal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'============================================================================
'           Copyright© 2002 Kevin Bronson. All rights reserved.
'============================================================================

Public Event Change(strErrorMessage As String)

Public WithEvents IP_Address As IP_Address
Attribute IP_Address.VB_VarHelpID = -1


'Locked
Public Property Get Locked() As Boolean
    Locked = Text1.Locked
End Property

'Locked
Public Property Let Locked(blnInput As Boolean)
    Text1.Locked = blnInput
    Text2.Locked = blnInput
    Text3.Locked = blnInput
    Text4.Locked = blnInput
End Property


'UpdateDisplay
Public Sub UpdateDisplay()
    Text1 = IP_Address.Octet1
    Text2 = IP_Address.Octet2
    Text3 = IP_Address.Octet3
    Text4 = IP_Address.Octet4
End Sub

'HandleChange
Private Sub HandleChange(objText As TextBox, lngOctetNumber As Long)
    Dim bytTest As Byte
    Dim strErrorMessage As String
    On Error GoTo Had_Error
    'Must do it this way to get IP_Address object's Change event to fire
    Select Case lngOctetNumber
        Case 1: IP_Address.Octet1 = CByte(objText)
        Case 2: IP_Address.Octet2 = CByte(objText)
        Case 3: IP_Address.Octet3 = CByte(objText)
        Case 4: IP_Address.Octet4 = CByte(objText)
    End Select
    GoTo Finalize_Sub
    
Had_Error:
    strErrorMessage = "'" & objText.Text & "' is NOT a Valid IP Address Octet."
    objText.SetFocus

Finalize_Sub:
    RaiseEvent Change(strErrorMessage)
End Sub

'IP_Address_Change
Private Sub IP_Address_Change()
    Call UpdateDisplay
End Sub

'Text1_KeyPress
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii < 58 And KeyAscii > 47) Or KeyAscii = 8 Or (KeyAscii = 3)) Then KeyAscii = 0
End Sub

'Text1_LostFocus
Private Sub Text1_LostFocus()
    Call HandleChange(Text1, 1)
End Sub

'Text2_KeyPress
Private Sub Text2_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii < 58 And KeyAscii > 47) Or KeyAscii = 8 Or (KeyAscii = 3)) Then KeyAscii = 0
End Sub

'Text2_LostFocus
Private Sub Text2_LostFocus()
    Call HandleChange(Text2, 2)
End Sub

'Text3_KeyPress
Private Sub Text3_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii < 58 And KeyAscii > 47) Or KeyAscii = 8 Or (KeyAscii = 3)) Then KeyAscii = 0
End Sub

'Text3_LostFocus
Private Sub Text3_LostFocus()
    Call HandleChange(Text3, 3)
End Sub

'Text4_KeyPress
Private Sub Text4_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii < 58 And KeyAscii > 47) Or KeyAscii = 8 Or (KeyAscii = 3)) Then KeyAscii = 0
End Sub

'Text4_LostFocus
Private Sub Text4_LostFocus()
    Call HandleChange(Text4, 4)
End Sub

'UserControl_Initialize
Private Sub UserControl_Initialize()
    Set IP_Address = New IP_Address
    Call UpdateDisplay
End Sub













