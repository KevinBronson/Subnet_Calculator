VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   7755
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   555
      Left            =   2700
      TabIndex        =   2
      Text            =   "100100111"
      Top             =   450
      Width           =   3795
   End
   Begin SubnetCalculator.LONG_Num LONG_Num1 
      Height          =   1095
      Left            =   810
      TabIndex        =   1
      Top             =   2070
      Width           =   4605
      _extentx        =   8123
      _extenty        =   1614
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'============================================================================
'           Copyright© 2002 Kevin Bronson. All rights reserved.
'============================================================================

Private Sub Command1_Click()
LONG_Num1.BinaryValue = Text1
End Sub
