Attribute VB_Name = "modFindIPs"

'******************************************************************
'Created By Verburgh Peter.
' 07-23-2001
' verburgh.peter@skynet.be
'-------------------------------------
'With this small application , you can detect the IP's installed on your computer,
'including subnet mask , BroadcastAddr..
'
'I've wrote this because i've a programm that uses the winsock control, but,
'if you have multiple ip's installed on your pc , you could get by using the Listen
' method the wrong ip ...
'Because Winsock.Localip => detects the default ip installed on your PC ,
' and in most of the cases it could be the LAN (nic) not the WAN (nic)
'So then you have to use the Bind function ,to bind to your right ip..
'but how do you know & find that ip ?
'you can find it now by this appl.. it check's in the api.. IP Table..
'******************************************************************


Const MAX_IP = 5 'To make a buffer... i dont think you have more than 5 ip on your pc..

Type IPINFO
     dwAddr As Long ' IP address
    dwIndex As Long ' interface index
    dwMask As Long ' subnet mask
    dwBCastAddr As Long ' broadcast address
    dwReasmSize As Long ' assembly size
    unused1 As Integer ' not currently used
    unused2 As Integer '; not currently used
End Type

Type MIB_IPADDRTABLE
    dEntrys As Long 'number of entries in the table
    mIPInfo(MAX_IP) As IPINFO 'array of IP address entries
End Type

Type IP_Array
    mBuffer As MIB_IPADDRTABLE
    BufferLen As Long
End Type

Public Declare Function GetIpAddrTable Lib "IPHlpApi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long

'converts a Long to a string
Public Function ConvertAddressToString(longAddr As Long) As String
    Dim myByte(3) As Byte
    Dim Cnt As Long
    CopyMemory myByte(0), longAddr, 4
    For Cnt = 0 To 3
        ConvertAddressToString = ConvertAddressToString + CStr(myByte(Cnt)) + "."
    Next Cnt
    ConvertAddressToString = Left$(ConvertAddressToString, Len(ConvertAddressToString) - 1)
End Function

Public Function GetIPs()
Dim Ret As Long, Tel As Long
Dim bBytes() As Byte
Dim Listing As MIB_IPADDRTABLE
Dim IPs() As String
    
'On Error GoTo END1
    GetIpAddrTable ByVal 0&, Ret, True

    If Ret <= 0 Then Exit Function
    ReDim bBytes(0 To Ret - 1) As Byte
    'retrieve the data
    GetIpAddrTable bBytes(0), Ret, False
      
    'Get the first 4 bytes to get the entry's.. ip installed
    CopyMemory Listing.dEntrys, bBytes(0), 4
    
    ReDim IPs(0 To Listing.dEntrys - 1) As String
    For Tel = 0 To Listing.dEntrys - 1
        'Copy whole structure to Listing..
        CopyMemory Listing.mIPInfo(Tel), bBytes(4 + (Tel * Len(Listing.mIPInfo(0)))), Len(Listing.mIPInfo(Tel))
        IPs(Tel) = ConvertAddressToString(Listing.mIPInfo(Tel).dwAddr)
    Next

    GetIPs = IPs
'Exit Function
'END1:
'MsgBox "ERROR"
End Function
