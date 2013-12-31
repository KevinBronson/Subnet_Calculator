Attribute VB_Name = "Common"
Option Explicit

'============================================================================
'           Copyright© 2002 Kevin Bronson. All rights reserved.
'============================================================================

Public objCompanyDisplay As CompanyDisplay




'SubnetRange
Function SubnetRange(objSubnetName As IP_Address, objSubnetMask As IP_Address, _
                     objIP_FirstHost As IP_Address, objIP_LastHost As IP_Address, _
                     lngHosts As Long) As Boolean
    Dim a As Long
    Dim lngFirstZero As Long
    Dim lngNumberOfZeros As Long
    Dim strTemp As String
    
    If objSubnetMask.IS_Valid_Subnet_Mask Then
        objIP_FirstHost.Octet1 = objSubnetName.Octet1
        objIP_FirstHost.Octet2 = objSubnetName.Octet2
        objIP_FirstHost.Octet3 = objSubnetName.Octet3
        objIP_FirstHost.Octet4 = objSubnetName.Octet4 + 1
        
        lngFirstZero = InStr(Replace(objSubnetMask.IP_Address_Binary, ".", ""), "0")
        If lngFirstZero = 0 Then lngFirstZero = 2
        lngNumberOfZeros = 33 - lngFirstZero
        lngHosts = (2 ^ lngNumberOfZeros) - 2
        
        strTemp = BinaryOperation(Replace(objSubnetName.IP_Address_Binary, ".", ""), _
                                  DecimalToBinary(lngHosts), BINARY_OR)
                                  
        objIP_LastHost.Octet1 = CByte(BinaryToDecimal(Mid(strTemp, 1, 8)))
        objIP_LastHost.Octet2 = CByte(BinaryToDecimal(Mid(strTemp, 9, 8)))
        objIP_LastHost.Octet3 = CByte(BinaryToDecimal(Mid(strTemp, 17, 8)))
        objIP_LastHost.Octet4 = CByte(BinaryToDecimal(Mid(strTemp, 25, 8)))
        
        SubnetRange = True
    Else
        MsgBox "Invalid Subnet Mask. Refer to Help for more information."
    End If
End Function





