VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IP_Address"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'============================================================================
'           Copyright© 2002 Kevin Bronson. All rights reserved.
'============================================================================

Dim bytOctet1 As Byte
Dim bytOctet2 As Byte
Dim bytOctet3 As Byte
Dim bytOctet4 As Byte

Public Event Change()
Attribute Change.VB_Description = "Fired when ever the value of an octect is changed."


'===================================================================================
'                             PUBLIC PROPERTIES
'===================================================================================

'Octet1
Public Property Get Octet1() As Byte
Attribute Octet1.VB_Description = "This is the first number in an IP Address. For example, x would be the first octect in the following IP Address: x.2.3.4"
    Octet1 = bytOctet1
End Property

Public Property Let Octet1(ByVal bytInput As Byte)
    bytOctet1 = bytInput
    RaiseEvent Change
End Property

'Octet2
Public Property Get Octet2() As Byte
Attribute Octet2.VB_Description = "This is the second number in an IP Address. For example, x would be the second octect in the following IP Address: 1.x.3.4"
    Octet2 = bytOctet2
End Property

Public Property Let Octet2(ByVal bytInput As Byte)
    bytOctet2 = bytInput
    RaiseEvent Change
End Property

'Octet3
Public Property Get Octet3() As Byte
Attribute Octet3.VB_Description = "This is the third number in an IP Address. For example, x would be the third octect in the following IP Address: 1.2.x.4"
    Octet3 = bytOctet3
End Property

Public Property Let Octet3(ByVal bytInput As Byte)
    bytOctet3 = bytInput
    RaiseEvent Change
End Property

'Octet4
Public Property Get Octet4() As Byte
Attribute Octet4.VB_Description = "This is the forth number in an IP Address. For example, x would be the forth octect in the following IP Address: 1.2.3.x"
    Octet4 = bytOctet4
End Property

Public Property Let Octet4(ByVal bytInput As Byte)
    bytOctet4 = bytInput
    RaiseEvent Change
End Property


'===================================================================================
'                                 READ-ONLY
'===================================================================================

'IP_Address_Decimal
Public Property Get IP_Address_Decimal() As String
Attribute IP_Address_Decimal.VB_Description = "Returns the IP Address in decimal format. Example: 192.56.23.2"
    IP_Address_Decimal = CStr(Octet1) & "." & CStr(Octet2) & "." & _
                         CStr(Octet3) & "." & CStr(Octet4)
End Property

'IP_Address_HEX
Public Property Get IP_Address_HEX() As String
Attribute IP_Address_HEX.VB_Description = "Returns the IP Address in Hexidecimal format. Example: AF.06.8B.03"
    IP_Address_HEX = Hex$(Octet1) & "." & Hex$(Octet2) & "." & _
                     Hex$(Octet3) & "." & Hex$(Octet4)
End Property

'IP_Address_Binary
Public Property Get IP_Address_Binary() As String
Attribute IP_Address_Binary.VB_Description = "Returns the IP Address in binary format. Example: 1101 1100.1010 0011.1100 0001.1110 0111"
    IP_Address_Binary = ByteToBinary(Octet1) & "." & ByteToBinary(Octet2) & "." & _
                        ByteToBinary(Octet3) & "." & ByteToBinary(Octet4)
End Property

'IS_Valid_Subnet_Mask
Public Property Get IS_Valid_Subnet_Mask() As Boolean
Attribute IS_Valid_Subnet_Mask.VB_Description = "Returns True if the current IP Address makes for a vaild subnet mask."
    Dim strTemp As String
    Dim lngPos As Long
    
    strTemp = Replace(IP_Address_Binary, ".", "")
    'See if first number is zero
    If Mid(strTemp, 1, 1) = "0" Then GoTo TEST_FAILED
    'Remove leading ones
    '123456789
    '111111000
    lngPos = InStr(strTemp, "0") 'lngPos=7
    strTemp = Right(strTemp, Len(strTemp) - lngPos + 1) 'length should be 3
    'See if there are any more ones left
    If InStr(strTemp, "1") Then GoTo TEST_FAILED
    'Make sure last two bits are zeros
    If Len(strTemp) > 1 Then
        If Right(strTemp, 2) <> "00" Then GoTo TEST_FAILED
    Else
        'Not enough zeros at end
        GoTo TEST_FAILED
    End If
    
    IS_Valid_Subnet_Mask = True
    Exit Property
    
TEST_FAILED:
    IS_Valid_Subnet_Mask = False

End Property

'Class
Public Property Get Class() As String
Attribute Class.VB_Description = "Returns the IP Address Class that the current IP Address is in. For example, 184.168.50.3 is a Class B IP Address."
    Select Case Octet1
        Case 1 To 127: Class = "A"
        'Case 127: Class = "A [Reserved for Loopback Testing]"
        Case 128 To 191: Class = "B"
        Case 192 To 223: Class = "C"
        Case 224 To 239: Class = "D"
        Case 240 To 254: Class = "E"
        Case Else: Class = "Unknown"
    End Select
End Property

'IsRoutable
Public Property Get IsRoutable() As Boolean
Attribute IsRoutable.VB_Description = "Returns True if the current IP Address is routable."
    IsRoutable = Not ((Octet1 = 10) Or _
                ((Octet1 = 172) And (Octet2 > 15 And Octet2 < 32)) Or _
                (Octet1 = 192 And Octet2 = 168))
End Property


'===================================================================================
'                              PUBLIC METHODS
'===================================================================================

'Reset
Public Sub Reset()
Attribute Reset.VB_Description = "Sets the IP Address to '0.0.0.0' and raises the change event."
    bytOctet1 = 0
    bytOctet2 = 0
    bytOctet3 = 0
    'Call the property itself so that the change event will be raised
    Octet4 = 0
End Sub


'===================================================================================
'                             PRIVATE FUNCTIONS
'===================================================================================

'ByteToBinary
Private Function ByteToBinary(ByVal lngByte As Byte) As String
    ByteToBinary = Hex(lngByte)
    If lngByte < 16 Then ByteToBinary = "0" & ByteToBinary
    ByteToBinary = HexToBinary(ByteToBinary)
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






