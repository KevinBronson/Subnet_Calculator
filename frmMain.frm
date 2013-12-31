VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Subnet Calculator"
   ClientHeight    =   7320
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9480
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmSubnetCalculations 
      BorderStyle     =   0  'None
      Height          =   6675
      Left            =   180
      TabIndex        =   30
      Top             =   450
      Width           =   9105
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1005
         Left            =   5940
         Picture         =   "frmMain.frx":0CCA
         ScaleHeight     =   975
         ScaleWidth      =   3045
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   90
         Width           =   3075
      End
      Begin VB.CommandButton cmdSubnet_Sample 
         Caption         =   "Use Sample Data"
         Height          =   375
         Left            =   3870
         TabIndex        =   29
         Top             =   2880
         Width           =   1815
      End
      Begin VB.CommandButton cmdSubnet_Reset 
         Caption         =   "Reset"
         Height          =   375
         Left            =   1980
         TabIndex        =   28
         Top             =   2880
         Width           =   1815
      End
      Begin VB.CommandButton cmdSubnet_Calc 
         Caption         =   "Calculate"
         Height          =   375
         Left            =   90
         TabIndex        =   27
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Frame frmResults 
         Caption         =   "Results"
         Height          =   3165
         Left            =   90
         TabIndex        =   0
         Top             =   3420
         Visible         =   0   'False
         Width           =   8925
         Begin VB.TextBox txtNumberOfHosts 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   7560
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   270
            Width           =   1185
         End
         Begin VB.Frame frmResults_HostRange 
            Caption         =   "Host Range"
            Height          =   1545
            Left            =   90
            TabIndex        =   32
            Top             =   1530
            Width           =   8745
            Begin SubnetCalculator.IP_BYTE_Num vpcResults_Host1 
               Height          =   1185
               Left            =   90
               TabIndex        =   33
               TabStop         =   0   'False
               Top             =   270
               Width           =   4065
               _extentx        =   7170
               _extenty        =   2090
            End
            Begin SubnetCalculator.IP_BYTE_Num vpcResults_Host2 
               Height          =   1185
               Left            =   4590
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   270
               Width           =   4065
               _extentx        =   7170
               _extenty        =   2090
            End
            Begin VB.Label Label1 
               Caption         =   ">>"
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
               Left            =   4230
               TabIndex        =   35
               Top             =   810
               Width           =   285
            End
         End
         Begin SubnetCalculator.IP_BYTE_Num vpcResults_NetworkID 
            Height          =   1185
            Left            =   180
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   270
            Width           =   4065
            _extentx        =   7170
            _extenty        =   2090
         End
         Begin VB.Label lblNumberOfHosts 
            Caption         =   "Host ID's available to this Network ID:"
            Height          =   195
            Left            =   4680
            TabIndex        =   38
            Top             =   360
            Width           =   2895
         End
      End
      Begin VB.Frame frmInputs 
         Caption         =   "Inputs"
         Height          =   1545
         Left            =   90
         TabIndex        =   20
         Top             =   1170
         Width           =   8925
         Begin VB.CommandButton cmdSubnetMaskC 
            Cancel          =   -1  'True
            Caption         =   "C"
            Height          =   285
            Left            =   8550
            TabIndex        =   26
            ToolTipText     =   "Default Class C Subnet Mask"
            Top             =   1170
            Width           =   285
         End
         Begin VB.CommandButton cmdSubnetMaskB 
            Caption         =   "B"
            Height          =   285
            Left            =   8550
            TabIndex        =   25
            ToolTipText     =   "Default Class B Subnet Mask"
            Top             =   900
            Width           =   285
         End
         Begin VB.CommandButton cmdSubnetMaskA 
            Caption         =   "A"
            Height          =   285
            Left            =   8550
            TabIndex        =   24
            ToolTipText     =   "Default Class A Subnet Mask"
            Top             =   630
            Width           =   285
         End
         Begin VB.CommandButton cmdSubnetMaskHelp 
            Caption         =   "?"
            Height          =   285
            Left            =   8550
            TabIndex        =   23
            ToolTipText     =   "Click here for a list of valid octets."
            Top             =   360
            Width           =   285
         End
         Begin SubnetCalculator.IP_BYTE_Num vpcInput_IP 
            Height          =   1185
            Left            =   180
            TabIndex        =   21
            Top             =   270
            Width           =   4065
            _extentx        =   7170
            _extenty        =   2090
         End
         Begin SubnetCalculator.IP_BYTE_Num vpcInput_SubnetMask 
            Height          =   1185
            Left            =   4410
            TabIndex        =   22
            Top             =   270
            Width           =   4065
            _extentx        =   7170
            _extenty        =   2090
         End
      End
   End
   Begin VB.Frame frmBinaryTools 
      BorderStyle     =   0  'None
      Height          =   6585
      Left            =   180
      TabIndex        =   36
      Top             =   450
      Width           =   9015
      Begin VB.Frame Frame1 
         Caption         =   "Variable Length Binary Numbers"
         Height          =   1725
         Left            =   90
         TabIndex        =   40
         Top             =   180
         Width           =   8925
         Begin VB.CommandButton cmdNAND 
            Caption         =   "NAND"
            Height          =   465
            Left            =   6660
            TabIndex        =   5
            Top             =   450
            Width           =   915
         End
         Begin VB.CommandButton cmdOR 
            Caption         =   "OR"
            Height          =   465
            Left            =   5580
            TabIndex        =   7
            Top             =   990
            Width           =   1005
         End
         Begin VB.CommandButton cmdNOR 
            Caption         =   "NOR"
            Height          =   465
            Left            =   6660
            TabIndex        =   8
            Top             =   990
            Width           =   915
         End
         Begin VB.CommandButton cmdAND 
            Caption         =   "AND"
            Height          =   465
            Left            =   5580
            TabIndex        =   4
            Top             =   450
            Width           =   1005
         End
         Begin VB.TextBox txtBIN3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   1170
            Width           =   3795
         End
         Begin VB.TextBox txtBIN2 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1440
            TabIndex        =   2
            Text            =   "1010"
            Top             =   810
            Width           =   3795
         End
         Begin VB.TextBox txtBIN1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1440
            TabIndex        =   1
            Text            =   "1100"
            Top             =   450
            Width           =   3795
         End
         Begin VB.CommandButton cmdIMPLIES 
            Caption         =   "IMPLIES"
            Height          =   465
            Left            =   7650
            TabIndex        =   6
            Top             =   450
            Width           =   1005
         End
         Begin VB.CommandButton cmdXOR 
            Caption         =   "XOR"
            Height          =   465
            Left            =   7650
            TabIndex        =   9
            Top             =   990
            Width           =   1005
         End
         Begin VB.Label Label5 
            Caption         =   "Answer:"
            Height          =   285
            Left            =   810
            TabIndex        =   43
            Top             =   1170
            Width           =   645
         End
         Begin VB.Label Label4 
            Caption         =   "Binary Number 2:"
            Height          =   285
            Left            =   180
            TabIndex        =   42
            Top             =   810
            Width           =   1275
         End
         Begin VB.Label Label3 
            Caption         =   "Binary Number 1:"
            Height          =   285
            Left            =   180
            TabIndex        =   41
            Top             =   450
            Width           =   1275
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "32 Bit Binary Numbers"
         Height          =   4605
         Left            =   90
         TabIndex        =   44
         Top             =   1980
         Width           =   8925
         Begin VB.CommandButton cmdNAND2 
            Caption         =   "NAND"
            Height          =   465
            Left            =   6660
            TabIndex        =   14
            Top             =   450
            Width           =   915
         End
         Begin VB.CommandButton cmdOR2 
            Caption         =   "OR"
            Height          =   465
            Left            =   5580
            TabIndex        =   16
            Top             =   990
            Width           =   1005
         End
         Begin VB.CommandButton cmdNOR2 
            Caption         =   "NOR"
            Height          =   465
            Left            =   6660
            TabIndex        =   17
            Top             =   990
            Width           =   915
         End
         Begin VB.CommandButton cmdAND2 
            Caption         =   "AND"
            Height          =   465
            Left            =   5580
            TabIndex        =   13
            Top             =   450
            Width           =   1005
         End
         Begin VB.CommandButton cmdImplies2 
            Caption         =   "IMPLIES"
            Height          =   465
            Left            =   7650
            TabIndex        =   15
            Top             =   450
            Width           =   1005
         End
         Begin VB.CommandButton cmdXOR2 
            Caption         =   "XOR"
            Height          =   465
            Left            =   7650
            TabIndex        =   18
            Top             =   990
            Width           =   1005
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1005
            Left            =   5580
            Picture         =   "frmMain.frx":294A
            ScaleHeight     =   975
            ScaleWidth      =   3045
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   3060
            Width           =   3075
         End
         Begin SubnetCalculator.LONG_Num LONG_Num1 
            Height          =   1095
            Left            =   1440
            TabIndex        =   10
            Top             =   450
            Width           =   3615
            _extentx        =   6376
            _extenty        =   1931
         End
         Begin SubnetCalculator.LONG_Num LONG_Num3 
            Height          =   1095
            Left            =   1440
            TabIndex        =   12
            Top             =   2970
            Width           =   3615
            _extentx        =   6376
            _extenty        =   1931
         End
         Begin SubnetCalculator.LONG_Num LONG_Num2 
            Height          =   1095
            Left            =   1440
            TabIndex        =   11
            Top             =   1710
            Width           =   3615
            _extentx        =   6376
            _extenty        =   1931
         End
         Begin VB.Label Label2 
            Caption         =   "Binary Number 1:"
            Height          =   285
            Left            =   180
            TabIndex        =   47
            Top             =   1260
            Width           =   1275
         End
         Begin VB.Label Label7 
            Caption         =   "Answer:"
            Height          =   285
            Left            =   810
            TabIndex        =   46
            Top             =   3780
            Width           =   645
         End
         Begin VB.Label Label6 
            Caption         =   "Binary Number 2:"
            Height          =   285
            Left            =   180
            TabIndex        =   45
            Top             =   2520
            Width           =   1275
         End
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   7125
      Left            =   90
      TabIndex        =   48
      Top             =   90
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   12568
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Subnet Calculator"
            Key             =   "SubnetCalculator"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Binary Functions"
            Key             =   "BinaryFunctions"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   180
      X2              =   9270
      Y1              =   5940
      Y2              =   5940
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu Spacer1 
         Caption         =   "-"
      End
      Begin VB.Menu Close 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      Begin VB.Menu HelpContents 
         Caption         =   "&Contents..."
      End
      Begin VB.Menu ValidSubnetMasks 
         Caption         =   "&Valid Subnet Masks"
      End
      Begin VB.Menu Spacer2 
         Caption         =   "-"
      End
      Begin VB.Menu LegalStatement 
         Caption         =   "&Legal Statement..."
      End
      Begin VB.Menu VersionInfo 
         Caption         =   "&Version Information..."
      End
      Begin VB.Menu Spacer3 
         Caption         =   "-"
      End
      Begin VB.Menu About 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'============================================================================
'           Copyright© 2002 Kevin Bronson. All rights reserved.
'============================================================================


'===================================================================================
'                                MENU ITEMS
'===================================================================================

'Close_Click
Private Sub Close_Click()
    End
End Sub

'About_Click
Private Sub About_Click()
    objCompanyDisplay.About
End Sub

'HelpContents_Click
Private Sub HelpContents_Click()
    Dim objShell As New Shell32.Shell
    objShell.Open (App.Path & "\Help\help.chm")
    Set objShell = Nothing
End Sub

'LegalStatement_Click
Private Sub LegalStatement_Click()
    'Trap Hard Errors
    On Error GoTo HadHardError
    
    frmLegalStatement.Show
    
    Exit Sub
HadHardError:
    MsgBox FormatErrorMessage(Err, Me.Name, "LegalStatement_Click")
End Sub

'TabStrip1_Click
Private Sub TabStrip1_Click()
    Select Case TabStrip1.SelectedItem.Key
        Case "SubnetCalculator"
            Call HideAllFrames
            frmSubnetCalculations.Visible = True
        Case "BinaryFunctions"
            Call HideAllFrames
            frmBinaryTools.Visible = True
    End Select
End Sub

'ValidSubnetMasks_Click
Private Sub ValidSubnetMasks_Click()
    MsgBox ValidSubnetMaskOctets
End Sub



'===================================================================================
'                            SUBNET CALCULATOR BUTTONS
'===================================================================================

'cmdSubnet_Calc_Click
Private Sub cmdSubnet_Calc_Click()
    Dim strNetworkID As String
    Dim lngNumberOfHosts As Long
    
    strNetworkID = BinaryOperation(Replace(vpcInput_IP.IP_Address.IP_Address_Binary, ".", ""), _
                   Replace(vpcInput_SubnetMask.IP_Address.IP_Address_Binary, ".", ""), BINARY_AND)
    
    vpcResults_NetworkID.IP_Address.Octet1 = CByte(BinaryToDecimal(Mid(strNetworkID, 1, 8)))
    vpcResults_NetworkID.IP_Address.Octet2 = CByte(BinaryToDecimal(Mid(strNetworkID, 9, 8)))
    vpcResults_NetworkID.IP_Address.Octet3 = CByte(BinaryToDecimal(Mid(strNetworkID, 17, 8)))
    vpcResults_NetworkID.IP_Address.Octet4 = CByte(BinaryToDecimal(Mid(strNetworkID, 25, 8)))
    
    If SubnetRange(vpcResults_NetworkID.IP_Address, _
                   vpcInput_SubnetMask.IP_Address, _
                   vpcResults_Host1.IP_Address, _
                   vpcResults_Host2.IP_Address, _
                   lngNumberOfHosts) Then frmResults.Visible = True
    
    txtNumberOfHosts = Format(lngNumberOfHosts, "#,##0")
End Sub

'cmdSubnet_Reset_Click
Private Sub cmdSubnet_Reset_Click()
    frmResults.Visible = False
    vpcInput_IP.Reset
    vpcInput_SubnetMask.Reset
    vpcResults_NetworkID.Reset
    vpcResults_Host1.Reset
    vpcResults_Host2.Reset
End Sub

'cmdSubnet_Sample_Click
Private Sub cmdSubnet_Sample_Click()
    vpcInput_IP.IP_Address.Octet1 = 172
    vpcInput_IP.IP_Address.Octet2 = 16
    vpcInput_IP.IP_Address.Octet3 = 10
    vpcInput_IP.IP_Address.Octet4 = 37
    vpcInput_SubnetMask.IP_Address.Octet1 = 255
    vpcInput_SubnetMask.IP_Address.Octet2 = 255
    vpcInput_SubnetMask.IP_Address.Octet3 = 128
    vpcInput_SubnetMask.IP_Address.Octet4 = 0
End Sub

'cmdSubnetMaskHelp_Click
Private Sub cmdSubnetMaskHelp_Click()
    MsgBox ValidSubnetMaskOctets
End Sub

'cmdSubnetMaskA_Click
Private Sub cmdSubnetMaskA_Click()
    vpcInput_SubnetMask.IP_Address.Octet1 = 255
    vpcInput_SubnetMask.IP_Address.Octet2 = 0
    vpcInput_SubnetMask.IP_Address.Octet3 = 0
    vpcInput_SubnetMask.IP_Address.Octet4 = 0
End Sub

'cmdSubnetMaskB_Click
Private Sub cmdSubnetMaskB_Click()
    vpcInput_SubnetMask.IP_Address.Octet1 = 255
    vpcInput_SubnetMask.IP_Address.Octet2 = 255
    vpcInput_SubnetMask.IP_Address.Octet3 = 0
    vpcInput_SubnetMask.IP_Address.Octet4 = 0
End Sub

'cmdSubnetMaskC_Click
Private Sub cmdSubnetMaskC_Click()
    vpcInput_SubnetMask.IP_Address.Octet1 = 255
    vpcInput_SubnetMask.IP_Address.Octet2 = 255
    vpcInput_SubnetMask.IP_Address.Octet3 = 255
    vpcInput_SubnetMask.IP_Address.Octet4 = 0
End Sub

'===================================================================================
'                            BINARY OPERATIONS BUTTONS
'===================================================================================

'cmdAND_Click
Private Sub cmdAND_Click()
    If txtBIN1 <> "" And txtBIN2 <> "" Then txtBIN3 = BinaryOperation(txtBIN1, txtBIN2, BINARY_AND)
End Sub

'cmdNAND_Click
Private Sub cmdNAND_Click()
    If txtBIN1 <> "" And txtBIN2 <> "" Then txtBIN3 = BinaryOperation(txtBIN1, txtBIN2, BINARY_NAND)
End Sub

'cmdNOR_Click
Private Sub cmdNOR_Click()
    If txtBIN1 <> "" And txtBIN2 <> "" Then txtBIN3 = BinaryOperation(txtBIN1, txtBIN2, BINARY_NOR)
End Sub

'cmdOR_Click
Private Sub cmdOR_Click()
    If txtBIN1 <> "" And txtBIN2 <> "" Then txtBIN3 = BinaryOperation(txtBIN1, txtBIN2, BINARY_OR)
End Sub

'cmdIMPLIES_Click
Private Sub cmdIMPLIES_Click()
    If txtBIN1 <> "" And txtBIN2 <> "" Then txtBIN3 = BinaryOperation(txtBIN1, txtBIN2, BINARY_IMPLIES)
End Sub

'cmdXOR_Click
Private Sub cmdXOR_Click()
    If txtBIN1 <> "" And txtBIN2 <> "" Then txtBIN3 = BinaryOperation(txtBIN1, txtBIN2, BINARY_XOR)
End Sub

'===================================================================================

'cmdAND2_Click
Private Sub cmdAND2_Click()
    LONG_Num3.BinaryValue = BinaryOperation(LONG_Num1.BinaryValue, LONG_Num2.BinaryValue, BINARY_AND)
End Sub

'cmdImplies2_Click
Private Sub cmdImplies2_Click()
    LONG_Num3.BinaryValue = BinaryOperation(LONG_Num1.BinaryValue, LONG_Num2.BinaryValue, BINARY_IMPLIES)
End Sub

'cmdNAND2_Click
Private Sub cmdNAND2_Click()
    LONG_Num3.BinaryValue = BinaryOperation(LONG_Num1.BinaryValue, LONG_Num2.BinaryValue, BINARY_NAND)
End Sub

'cmdNOR2_Click
Private Sub cmdNOR2_Click()
    LONG_Num3.BinaryValue = BinaryOperation(LONG_Num1.BinaryValue, LONG_Num2.BinaryValue, BINARY_NOR)
End Sub

'cmdOR2_Click
Private Sub cmdOR2_Click()
    LONG_Num3.BinaryValue = BinaryOperation(LONG_Num1.BinaryValue, LONG_Num2.BinaryValue, BINARY_OR)
End Sub

'cmdXOR2_Click
Private Sub cmdXOR2_Click()
    LONG_Num3.BinaryValue = BinaryOperation(LONG_Num1.BinaryValue, LONG_Num2.BinaryValue, BINARY_XOR)
End Sub


'===================================================================================
'                                 EVENTS
'===================================================================================

'Form_Load
Private Sub Form_Load()
    Set Picture1.Picture = objCompanyDisplay.LogoImage
    Set Picture2.Picture = objCompanyDisplay.LogoImage
    About.Caption = "About " & objCompanyDisplay.CompanyName & "..."
    vpcInput_IP.Caption = "IP Address"
    vpcInput_IP.ShowClassInCaption = True
    vpcInput_IP.ShowRoutableInCaption = True
    vpcInput_IP.UseAllTabStops = False
    vpcInput_SubnetMask.Caption = "Subnet Mask"
    vpcInput_SubnetMask.UseAllTabStops = False
    vpcResults_NetworkID.Caption = "Network ID"
    vpcResults_NetworkID.Locked = True
    vpcResults_Host1.Caption = "First Host"
    vpcResults_Host1.Locked = True
    vpcResults_Host2.Caption = "Last Host"
    vpcResults_Host2.Locked = True
    LONG_Num1.UseAllTabStops = False
    LONG_Num2.UseAllTabStops = False
    LONG_Num3.UseAllTabStops = False
    LONG_Num3.Locked = True
    TabStrip1.TabIndex = 1
    Call TabStrip1_Click
End Sub

'VersionInfo_Click
Private Sub VersionInfo_Click()
    MsgBox App.Title & " Version: " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

'vpcInput_SubnetMask_Validate
Private Sub vpcInput_SubnetMask_Validate(Cancel As Boolean)
    If Not (vpcInput_SubnetMask.IP_Address.IS_Valid_Subnet_Mask) Then
        If Not (vpcInput_SubnetMask.IP_Address.IP_Address_Decimal = "0.0.0.0") Then
            MsgBox "Invalid Subnet Mask. Refer to Help for more inforamtion."
        End If
    End If
End Sub


'Picture1_Click
Private Sub Picture1_Click()
    Call About_Click
End Sub

'Picture2_Click
Private Sub Picture2_Click()
    Call About_Click
End Sub

'txtBIN1_KeyPress
Private Sub txtBIN1_KeyPress(KeyAscii As Integer)
    If Not (IS_BinaryInput(KeyAscii)) Then KeyAscii = 0 'Cancels Input
End Sub

'txtBIN2_KeyPress
Private Sub txtBIN2_KeyPress(KeyAscii As Integer)
    If Not (IS_BinaryInput(KeyAscii)) Then KeyAscii = 0 'Cancels Input
End Sub


'===================================================================================
'                                 FUNCTIONS
'===================================================================================

'HideAllFrames
Private Sub HideAllFrames()
    frmBinaryTools.Visible = False
    frmSubnetCalculations.Visible = False
End Sub

'ValidSubnetMaskOctets_Click
Private Function ValidSubnetMaskOctets() As String
    Dim arrTemp(100) As String
    
    arrTemp(0) = "Valid Subnet Mask Octets  " & vbCrLf
    arrTemp(1) = "----------------------------" & vbCrLf
    arrTemp(2) = "BINARY        DECIMAL" & vbCrLf
    arrTemp(3) = "0000 0000        0   " & vbCrLf
    arrTemp(4) = "1000 0000       128  " & vbCrLf
    arrTemp(5) = "1100 0000       192  " & vbCrLf
    arrTemp(6) = "1110 0000       224  " & vbCrLf
    arrTemp(7) = "1111 0000       240  " & vbCrLf
    arrTemp(8) = "1111 1000       248  " & vbCrLf
    arrTemp(9) = "1111 1100       252  " & vbCrLf
    arrTemp(10) = "1111 1110       254  " & vbCrLf
    arrTemp(11) = "1111 1111       255  " & vbCrLf
    arrTemp(12) = "----------------------------" & vbCrLf
    arrTemp(13) = "NOTE: " & vbCrLf
    arrTemp(14) = "  First Octet can not be 0" & vbCrLf
    arrTemp(15) = "  Very last octet can not be 255" & vbCrLf
    arrTemp(16) = "  Very last octet should not be 254" & vbCrLf
    
    ValidSubnetMaskOctets = Join(arrTemp, "")
End Function

'IS_BinaryInput
Private Function IS_BinaryInput(KeyAscii As Integer)
    'KeyCode '96' is number 1 from key pad
    'KeyCode '97' is number 0 from key pad
    IS_BinaryInput = (KeyAscii = Asc("0") Or KeyAscii = Asc("1") Or _
                      KeyAscii = 96 Or KeyAscii = 97 Or KeyAscii = 8)
End Function
















