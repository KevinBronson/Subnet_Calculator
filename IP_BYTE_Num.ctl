VERSION 5.00
Begin VB.UserControl IP_BYTE_Num 
   ClientHeight    =   1545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4470
   ScaleHeight     =   1545
   ScaleWidth      =   4470
   Begin VB.Frame Frame1 
      Caption         =   "0.0.0.0"
      Height          =   1185
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4065
      Begin SubnetCalculator.BYTE_Num BYTE_Num1 
         Height          =   825
         Left            =   90
         TabIndex        =   1
         Top             =   270
         Width           =   915
         _extentx        =   1614
         _extenty        =   1455
      End
      Begin SubnetCalculator.BYTE_Num BYTE_Num2 
         Height          =   825
         Left            =   1080
         TabIndex        =   2
         Top             =   270
         Width           =   915
         _extentx        =   1614
         _extenty        =   1455
      End
      Begin SubnetCalculator.BYTE_Num BYTE_Num3 
         Height          =   825
         Left            =   2070
         TabIndex        =   3
         Top             =   270
         Width           =   915
         _extentx        =   1614
         _extenty        =   1455
      End
      Begin SubnetCalculator.BYTE_Num BYTE_Num4 
         Height          =   825
         Left            =   3060
         TabIndex        =   4
         Top             =   270
         Width           =   915
         _extentx        =   1614
         _extenty        =   1455
      End
   End
End
Attribute VB_Name = "IP_BYTE_Num"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'============================================================================
'           Copyright© 2002 Kevin Bronson. All rights reserved.
'============================================================================

'===================================================================================
'                             DELCARATIONS
'===================================================================================

Public WithEvents IP_Address As IP_Address
Attribute IP_Address.VB_VarHelpID = -1

Public Event Validate(blnCancel As Boolean)
Public Event Change(blnCancel As Boolean, ByVal bytNewValue As Byte, ByVal bytOldValue As Byte)
Public Event UserChange(blnCancel As Boolean, ByVal bytNewValue As Byte, ByVal bytOldValue As Byte)
Public Event Error(strErrorMessage As String)

Public ShowClassInCaption As Boolean
Public ShowRoutableInCaption As Boolean

Dim strCaption As String


'===================================================================================
'                             PUBLIC PROPERTIES
'===================================================================================

'Caption
Public Property Get Caption() As String
    Caption = strCaption
End Property

Public Property Let Caption(strInput As String)
    strCaption = strInput
    Call UpdateFrameCaption
End Property

'UseAllTabStops
Public Property Get UseAllTabStops() As Boolean
    UseAllTabStops = BYTE_Num1.UseAllTabStops
End Property

Public Property Let UseAllTabStops(blnInput As Boolean)
    BYTE_Num1.UseAllTabStops = blnInput
    BYTE_Num2.UseAllTabStops = blnInput
    BYTE_Num3.UseAllTabStops = blnInput
    BYTE_Num4.UseAllTabStops = blnInput
End Property

'Locked
Public Property Get Locked() As Boolean
    Locked = BYTE_Num1.Locked
End Property

Public Property Let Locked(blnInput As Boolean)
    BYTE_Num1.Locked = blnInput
    BYTE_Num2.Locked = blnInput
    BYTE_Num3.Locked = blnInput
    BYTE_Num4.Locked = blnInput
End Property

'Enabled
Public Property Get Enabled() As Boolean
    Enabled = BYTE_Num1.Enabled
End Property

Public Property Let Enabled(blnInput As Boolean)
    BYTE_Num1.Enabled = blnInput
    BYTE_Num2.Enabled = blnInput
    BYTE_Num3.Enabled = blnInput
    BYTE_Num4.Enabled = blnInput
End Property




'===================================================================================
'                             INTERNAL EVENTS
'===================================================================================

'IP_Address_Change
Private Sub IP_Address_Change()
    BYTE_Num1.Value = IP_Address.Octet1
    BYTE_Num2.Value = IP_Address.Octet2
    BYTE_Num3.Value = IP_Address.Octet3
    BYTE_Num4.Value = IP_Address.Octet4
    Call UpdateFrameCaption
End Sub

'BYTE_Num1_UserChange
Private Sub BYTE_Num1_UserChange(blnCancel As Boolean, ByVal bytNewValue As Byte, ByVal bytOldValue As Byte)
    IP_Address.Octet1 = BYTE_Num1.Value
End Sub

'BYTE_Num2_UserChange
Private Sub BYTE_Num2_UserChange(blnCancel As Boolean, ByVal bytNewValue As Byte, ByVal bytOldValue As Byte)
    IP_Address.Octet2 = BYTE_Num2.Value
End Sub

'BYTE_Num3_UserChange
Private Sub BYTE_Num3_UserChange(blnCancel As Boolean, ByVal bytNewValue As Byte, ByVal bytOldValue As Byte)
    IP_Address.Octet3 = BYTE_Num3.Value
End Sub

'BYTE_Num4_UserChange
Private Sub BYTE_Num4_UserChange(blnCancel As Boolean, ByVal bytNewValue As Byte, ByVal bytOldValue As Byte)
    IP_Address.Octet4 = BYTE_Num4.Value
End Sub

'BYTE_Num1_Validate
Private Sub BYTE_Num1_Validate(Cancel As Boolean)
    RaiseEvent Validate(Cancel)
End Sub

'BYTE_Num2_Validate
Private Sub BYTE_Num2_Validate(Cancel As Boolean)
    RaiseEvent Validate(Cancel)
End Sub

'BYTE_Num3_Validate
Private Sub BYTE_Num3_Validate(Cancel As Boolean)
    RaiseEvent Validate(Cancel)
End Sub

'BYTE_Num4_Validate
Private Sub BYTE_Num4_Validate(Cancel As Boolean)
    RaiseEvent Validate(Cancel)
End Sub

'UserControl_Initialize
Private Sub UserControl_Initialize()
    Set IP_Address = New IP_Address
    strCaption = "Caption"
    Call UpdateFrameCaption
End Sub


'===================================================================================
'                             PUBLIC METHODS
'===================================================================================

'Reset
Public Sub Reset()
    IP_Address.Reset
End Sub


'===================================================================================
'                             PRIVATE FUNCTIONS
'===================================================================================

'UpdateFrameCaption
Private Sub UpdateFrameCaption()
    If Len(Caption) > 0 Then
        Frame1.Caption = Caption & " - " & IP_Address.IP_Address_Decimal
        If ShowRoutableInCaption Or ShowClassInCaption Then
            Frame1.Caption = Frame1.Caption & " ("
        End If
        If ShowClassInCaption Then
            Frame1.Caption = Frame1.Caption & "Class " & IP_Address.Class
        End If
        If ShowRoutableInCaption And ShowClassInCaption Then
            Frame1.Caption = Frame1.Caption & ", "
        End If
        If ShowRoutableInCaption Then
            If IP_Address.IsRoutable Then
                Frame1.Caption = Frame1.Caption & "Routable"
            Else
                Frame1.Caption = Frame1.Caption & "Non-Routable"
            End If
        End If
        If ShowRoutableInCaption Or ShowClassInCaption Then
            Frame1.Caption = Frame1.Caption & ")"
        End If
    Else
        Frame1.Caption = IP_Address.IP_Address_Decimal
    End If
End Sub






















