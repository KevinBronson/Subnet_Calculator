VERSION 5.00
Begin VB.UserControl Octet 
   ClientHeight    =   570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1395
   ScaleHeight     =   570
   ScaleWidth      =   1395
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   450
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "0000"
      Top             =   0
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "0000"
      Top             =   0
      Width           =   465
   End
End
Attribute VB_Name = "Octet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'============================================================================
'           Copyright© 2002 Kevin Bronson. All rights reserved.
'============================================================================

Dim lngValue As Byte


'Value
Public Property Get Value() As Byte
    Value = lngValue
End Property

Public Property Let Value(lngInput As Byte)
    lngValue = lngInput
    Call UpdateDisplay
End Property

'UpdateDisplay
Public Sub UpdateDisplay()
    Text1 = Left(ByteToBinary(lngValue), 4)
    Text2 = Right(ByteToBinary(lngValue), 4)
End Sub

'Text1_KeyPress
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii = Asc("1")) Or (KeyAscii = Asc("0")) Or (KeyAscii = 8) Or (KeyAscii = 3)) Then KeyAscii = 0
End Sub

'Text2_KeyPress
Private Sub Text2_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii = Asc("1")) Or (KeyAscii = Asc("0")) Or (KeyAscii = 8) Or (KeyAscii = 3)) Then KeyAscii = 0
End Sub

'UserControl_Initialize
Private Sub UserControl_Initialize()
    Call UpdateDisplay
End Sub
