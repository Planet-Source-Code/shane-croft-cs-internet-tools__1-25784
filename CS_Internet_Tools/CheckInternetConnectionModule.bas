Attribute VB_Name = "CheckInternetConnectionModule"
'**************************************
' Name: Check internet connection type
' This code makes API calls to check for an internet conection and
' returns the type of connection through API calls to the wininet.dll
' By: Ayman Elbanhawy
'
' Inputs: None
'
' Returns: None
'
' Assumes: None
'
'Side Effects: None
'This code is copyrighted and has limited warranties.
'Please visit our site for an updated version of this code.
' http://www.imagineer-web.com/MasterKey/VB/connectionCheck.htm
'**************************************
'START .BAS MODULE CODE
Option Explicit

Public Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Long
    'Internet connection VIA Proxy server.
    Public Const ProxyConnection As Long = &H4
    
    'Modem is busy.
    Public Const ModemConnectionIsBusy As Long = &H8

    'Internet connection is currently Offline
    Public Const InternetIsOffline As Long = &H20
    
    'Internet connection is currently configured
    Public Const InternetConnectionIsConfigured As Long = &H40
    
    'Internet connection VIA Modem.
    Public Const ModemConnection As Long = &H1
    
    'Remote Access Server is installed.
    Public Const RasInstalled As Long = &H10

    'Internet connection VIA LAN.
    Public Const LanConnection As Long = &H2
    
    
    


Public Function IsLanConnection() As Boolean
    Dim dwflags As Long
    'return True if LAN connection
    Call InternetGetConnectedState(dwflags, 0&)
    IsLanConnection = dwflags And LanConnection
End Function


Public Function IsModemConnection() As Boolean
    Dim dwflags As Long
    'return True if modem connection.
    Call InternetGetConnectedState(dwflags, 0&)
    IsModemConnection = dwflags And ModemConnection
End Function


Public Function IsProxyConnection() As Boolean
    Dim dwflags As Long
    'return True if connected through a proxy.
    Call InternetGetConnectedState(dwflags, 0&)
    IsProxyConnection = dwflags And ProxyConnection
End Function


Public Function IsConnected() As Boolean
    'Returns true if there is any internet connection.
    IsConnected = InternetGetConnectedState(0&, 0&)
End Function


Public Function IsRasInstalled() As Boolean
    Dim dwflags As Long
    'return True if RAS installed.
    Call InternetGetConnectedState(dwflags, 0&)
    IsRasInstalled = dwflags And RasInstalled
End Function


Public Function ConnectionTypeMsg(list As listbox) As String
    Dim dwflags As Long
    Dim msg As String
    'Return Internet connection msg.

    If InternetGetConnectedState(dwflags, 0&) Then

        If dwflags And InternetConnectionIsConfigured Then
            list.AddItem "Internet connection is configured." & vbCrLf
        End If

        If dwflags And LanConnection Then
            list.AddItem "Internet connection via a LAN"
        End If

        If dwflags And ProxyConnection Then
            list.AddItem "Connection is through a proxy server."
        End If

        If dwflags And ModemConnection Then
            list.AddItem "Internet connection via a Modem"
        End If

        If dwflags And InternetIsOffline Then
            list.AddItem "Internet connection is currently offline."
        End If

        If dwflags And ModemConnectionIsBusy Then
            list.AddItem "Modem is busy with a non-Internet connection."
        End If

        If dwflags And RasInstalled Then
            list.AddItem "Remote Access Services are installed on local system."
        End If
        
    Else
        list.AddItem "You are currently not connected to the internet."
        
    End If
    
    'ConnectionTypeMsg = msg
    
End Function
' END MODULE CODE
'##############################



                

