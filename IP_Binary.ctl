VERSION 5.00
Begin VB.UserControl IP_Binary 
   ClientHeight    =   780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4380
   ScaleHeight     =   780
   ScaleWidth      =   4380
   Begin SubnetCalculator.Octet Octet1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   503
   End
   Begin SubnetCalculator.Octet Octet2 
      Height          =   285
      Left            =   990
      TabIndex        =   1
      Top             =   0
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   503
   End
   Begin SubnetCalculator.Octet Octet3 
      Height          =   285
      Left            =   1980
      TabIndex        =   2
      Top             =   0
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   503
   End
   Begin SubnetCalculator.Octet Octet4 
      Height          =   285
      Left            =   2970
      TabIndex        =   3
      Top             =   0
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   503
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
      Left            =   2880
      TabIndex        =   6
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
      Left            =   1890
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
      Left            =   900
      TabIndex        =   4
      Top             =   0
      Width           =   105
   End
End
Attribute VB_Name = "IP_Binary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'============================================================================
'           Copyright© 2002 Kevin Bronson. All rights reserved.
'============================================================================

Public WithEvents IP_Address As IP_Address
Attribute IP_Address.VB_VarHelpID = -1


'UpdateDisplay
Public Sub UpdateDisplay()
    Octet1.Value = IP_Address.Octet1
    Octet2.Value = IP_Address.Octet2
    Octet3.Value = IP_Address.Octet3
    Octet4.Value = IP_Address.Octet4
End Sub

'IP_Address_Change
Private Sub IP_Address_Change()
    Call UpdateDisplay
End Sub


'UserControl_Initialize
Private Sub UserControl_Initialize()
    Set IP_Address = New IP_Address
    Call UpdateDisplay
End Sub





