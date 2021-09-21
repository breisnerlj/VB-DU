Attribute VB_Name = "apis"
Option Explicit

'Modified Martinnets-Resalta Objeto
Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'***********
Public Declare Function timeGetTime Lib "winmm.dll" () As Long

Public Const MAX_HOSTNAME_LEN = 132
Public Const MAX_DOMAIN_NAME_LEN = 132
Public Const MAX_SCOPE_ID_LEN = 260
Public Const MAX_ADAPTER_NAME_LENGTH = 260
Public Const MAX_ADAPTER_ADDRESS_LENGTH = 8
Public Const MAX_ADAPTER_DESCRIPTION_LENGTH = 132
Public Const ERROR_BUFFER_OVERFLOW = 111
Public Const MIB_IF_TYPE_ETHERNET = 6
Public Const MIB_IF_TYPE_TOKENRING = 9
Public Const MIB_IF_TYPE_FDDI = 15
Public Const MIB_IF_TYPE_PPP = 23
Public Const MIB_IF_TYPE_LOOPBACK = 24
Public Const MIB_IF_TYPE_SLIP = 28

Type IP_ADDR_STRING
            Next As Long
            IpAddress As String * 16
            IpMask As String * 16
            Context As Long
End Type

Type IP_ADAPTER_INFO
            Next As Long
            ComboIndex As Long
            AdapterName As String * MAX_ADAPTER_NAME_LENGTH
            Description As String * MAX_ADAPTER_DESCRIPTION_LENGTH
            AddressLength As Long
            address(MAX_ADAPTER_ADDRESS_LENGTH - 1) As Byte
            Index As Long
            Type As Long
            DhcpEnabled As Long
            CurrentIpAddress As Long
            IpAddressList As IP_ADDR_STRING
            GatewayList As IP_ADDR_STRING
            DhcpServer As IP_ADDR_STRING
            HaveWins As Byte
            PrimaryWinsServer As IP_ADDR_STRING
            SecondaryWinsServer As IP_ADDR_STRING
            LeaseObtained As Long
            LeaseExpires As Long
End Type

Type FIXED_INFO
            HostName As String * MAX_HOSTNAME_LEN
            DomainName As String * MAX_DOMAIN_NAME_LEN
            CurrentDnsServer As Long
            DnsServerList As IP_ADDR_STRING
            NodeType As Long
            ScopeId  As String * MAX_SCOPE_ID_LEN
            EnableRouting As Long
            EnableProxy As Long
            EnableDns As Long
End Type

Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Type RECTL
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Public Type SIZEL
    cx As Long
    cy As Long
End Type

Public Type FORM_INFO_1
    flags As Long
    pName As Long ' String
    Size As SIZEL
    ImageableArea As RECTL
End Type

Public Type sFORM_INFO_1
    flags As Long
    pName As String
    Size As SIZEL
    ImageableArea As RECTL
End Type


Public Declare Function EnumForms Lib "winspool.drv" Alias "EnumFormsA" (ByVal hPrinter As Long, ByVal level As Long, ByRef pForm As Any, ByVal cbBuf As Long, ByRef pcbNeeded As Long, ByRef pcReturned As Long) As Long

Public Declare Function AddForm Lib "winspool.drv" Alias "AddFormA" (ByVal hPrinter As Long, ByVal level As Long, pForm As Byte) As Long

Public Declare Function DeleteForm Lib "winspool.drv" Alias "DeleteFormA" (ByVal hPrinter As Long, ByVal pFormName As String) As Long

Public Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, ByVal pDefault As Long) As Long

Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long


Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByRef lpString2 As Long) As Long


Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (LpVersionInformation As OSVERSIONINFO) As Long


Public Declare Function GetNetworkParams Lib "IPHlpApi.dll" (FixedInfo As Any, pOutBufLen As Long) As Long
Public Declare Function GetAdaptersInfo Lib "IPHlpApi.dll" (IpAdapterInfo As Any, pOutBufLen As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
      End Type

Private Declare Function SetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME) As Long

'I.ECASTILLO 06.07.2021
Public Type POINTAPI
  X As Long
  Y As Long
End Type
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'F.ECASTILLO 06.07.2021
       
       

Function Obtener() As Variant
    Dim Error As Long
    Dim arrValores(22) As Variant
    Dim FixedInfoSize As Long
    Dim AdapterInfoSize As Long
    Dim i As Integer
    Dim PhysicalAddress  As String
    Dim NewTime As Date
    Dim AdapterInfo As IP_ADAPTER_INFO
    Dim AddrStr As IP_ADDR_STRING
    Dim FixedInfo As FIXED_INFO
    Dim buffer As IP_ADDR_STRING
    Dim pAddrStr As Long
    Dim pAdapt As Long
    Dim Buffer2 As IP_ADAPTER_INFO
    Dim FixedInfoBuffer() As Byte
    Dim AdapterInfoBuffer() As Byte
    
    ' Get the main IP configuration information for this machine
    ' using a FIXED_INFO structure.
    FixedInfoSize = 0
    Error = GetNetworkParams(ByVal 0&, FixedInfoSize)
    If Error <> 0 Then
        If Error <> ERROR_BUFFER_OVERFLOW Then
           MsgBox "GetNetworkParams sizing failed with error " & Error
           Exit Function
        End If
    End If
    ReDim FixedInfoBuffer(FixedInfoSize - 1)
    
    Error = GetNetworkParams(FixedInfoBuffer(0), FixedInfoSize)
    If Error = 0 Then
            CopyMemory FixedInfo, FixedInfoBuffer(0), FixedInfoSize
            arrValores(0) = FixedInfo.HostName 'Host Name
            arrValores(1) = FixedInfo.DnsServerList.IpAddress 'DNS Servers
            pAddrStr = FixedInfo.DnsServerList.Next
            Do While pAddrStr <> 0
                  CopyMemory buffer, ByVal pAddrStr, LenB(buffer)
                  arrValores(2) = buffer.IpAddress 'DNS Servers
                  pAddrStr = buffer.Next
            Loop
            
            Select Case FixedInfo.NodeType
                       Case 1
                                  arrValores(3) = "Node type: Broadcast"
                       Case 2
                                  arrValores(3) = "Node type: Peer to peer"
                       Case 4
                                  arrValores(3) = "Node type: Mixed"
                       Case 8
                                  arrValores(3) = "Node type: Hybrid"
                       Case Else
                                  arrValores(3) = "Unknown node type"
            End Select
            
            arrValores(4) = FixedInfo.ScopeId 'NetBIOS Scope ID
            If FixedInfo.EnableRouting Then
                       arrValores(5) = "IP Routing Enabled "
            Else
                       arrValores(5) = "IP Routing not enabled"
            End If
            If FixedInfo.EnableProxy Then
                       arrValores(6) = "WINS Proxy Enabled "
            Else
                       arrValores(6) = "WINS Proxy not Enabled "
            End If
            If FixedInfo.EnableDns Then
                      arrValores(7) = "NetBIOS Resolution Uses DNS "
            Else
                      arrValores(7) = "NetBIOS Resolution Does not use DNS  "
            End If
    Else
            MsgBox "GetNetworkParams failed with error " & Error
            Exit Function
    End If
    
    ' Enumerate all of the adapter specific information using the
    ' IP_ADAPTER_INFO structure.
    ' Note:  IP_ADAPTER_INFO contains a linked list of adapter entries.
    
    AdapterInfoSize = 0
    Error = GetAdaptersInfo(ByVal 0&, AdapterInfoSize)
    If Error <> 0 Then
        If Error <> ERROR_BUFFER_OVERFLOW Then
           MsgBox "GetAdaptersInfo sizing failed with error " & Error
           Exit Function
        End If
    End If
    ReDim AdapterInfoBuffer(AdapterInfoSize - 1)

    ' Get actual adapter information
    Error = GetAdaptersInfo(AdapterInfoBuffer(0), AdapterInfoSize)
    If Error <> 0 Then
       MsgBox "GetAdaptersInfo failed with error " & Error
       Exit Function
    End If
   
    ' Allocate memory
     CopyMemory AdapterInfo, AdapterInfoBuffer(0), AdapterInfoSize
    pAdapt = AdapterInfo.Next

    Do
     CopyMemory Buffer2, AdapterInfo, AdapterInfoSize
       Select Case Buffer2.Type
              Case MIB_IF_TYPE_ETHERNET
                  arrValores(8) = "Adapter name: Ethernet adapter "
              Case MIB_IF_TYPE_TOKENRING
                  arrValores(8) = "Adapter name: Token Ring adapter "
              Case MIB_IF_TYPE_FDDI
                    arrValores(8) = "Adapter name: FDDI adapter "
              Case MIB_IF_TYPE_PPP
                   arrValores(8) = "Adapter name: PPP adapter"
              Case MIB_IF_TYPE_LOOPBACK
                  arrValores(8) = "Adapter name: Loopback adapter "
              Case MIB_IF_TYPE_SLIP
                   arrValores(8) = "Adapter name: Slip adapter "
              Case Else
                    arrValores(8) = "Adapter name: Other adapter "
       End Select
       arrValores(9) = "AdapterDescription: " & Buffer2.Description

       PhysicalAddress = ""
       For i = 0 To Buffer2.AddressLength - 1
           PhysicalAddress = PhysicalAddress & Hex(Buffer2.address(i))
           If i < Buffer2.AddressLength - 1 Then
              PhysicalAddress = PhysicalAddress & "-"
           End If
       Next
       arrValores(10) = "Physical Address: " & PhysicalAddress
    
       If Buffer2.DhcpEnabled Then
          arrValores(11) = "DHCP Enabled "
       Else
          arrValores(11) = "DHCP disabled"
       End If

       arrValores(12) = Buffer2.IpAddressList.IpAddress 'ip de la maquin
       arrValores(13) = "Subnet Mask: " & Buffer2.IpAddressList.IpMask
       pAddrStr = Buffer2.IpAddressList.Next
       Do While pAddrStr <> 0
          CopyMemory buffer, Buffer2.IpAddressList, LenB(buffer)
          arrValores(14) = "IP Address: " & buffer.IpAddress
          arrValores(15) = "Subnet Mask: " & buffer.IpMask
          pAddrStr = buffer.Next
          If pAddrStr <> 0 Then
             CopyMemory Buffer2.IpAddressList, ByVal pAddrStr, _
                        LenB(Buffer2.IpAddressList)
          End If
       Loop
    
       arrValores(16) = "Default Gateway: " & Buffer2.GatewayList.IpAddress
       pAddrStr = Buffer2.GatewayList.Next
       Do While pAddrStr <> 0
          CopyMemory buffer, Buffer2.GatewayList, LenB(buffer)
          arrValores(17) = "IP Address: " & buffer.IpAddress
          pAddrStr = buffer.Next
          If pAddrStr <> 0 Then
             CopyMemory Buffer2.GatewayList, ByVal pAddrStr, _
                        LenB(Buffer2.GatewayList)
          End If
       Loop

       arrValores(18) = "DHCP Server: " & Buffer2.DhcpServer.IpAddress
       arrValores(19) = "Primary WINS Server: " & _
              Buffer2.PrimaryWinsServer.IpAddress
       arrValores(20) = "Secondary WINS Server: " & _
              Buffer2.SecondaryWinsServer.IpAddress

       ' Display time.
       NewTime = DateAdd("s", Buffer2.LeaseObtained, #1/1/1970#)
       arrValores(21) = "Lease Obtained: " & _
              CStr(Format(NewTime, "dddd, mmm d hh:mm:ss yyyy"))
     
       NewTime = DateAdd("s", Buffer2.LeaseExpires, #1/1/1970#)
       arrValores(22) = "Lease Expires :  " & _
              CStr(Format(NewTime, "dddd, mmm d hh:mm:ss yyyy"))
       pAdapt = Buffer2.Next
       If pAdapt <> 0 Then
           CopyMemory AdapterInfo, ByVal pAdapt, AdapterInfoSize
        End If
      Loop Until pAdapt = 0
      Obtener = arrValores()
End Function


Public Function GetFormName(ByVal PrinterHandle As Long, _
                            FormSize As SIZEL, FormName As String) As Integer
    Dim NumForms As Long, i As Long
    Dim FI1 As FORM_INFO_1
    Dim aFI1() As FORM_INFO_1 ' Working FI1 array
    Dim temp() As Byte ' Temp FI1 array
    Dim FormIndex As Integer
    Dim BytesNeeded As Long
    Dim RetVal As Long
    
    FormName = vbNullString
    FormIndex = 0
    ReDim aFI1(1)
    ' First call retrieves the BytesNeeded.
    RetVal = EnumForms(PrinterHandle, 1, aFI1(0), 0&, BytesNeeded, NumForms)
    ReDim temp(BytesNeeded)
    ReDim aFI1(BytesNeeded / Len(FI1))
    ' Second call actually enumerates the supported forms.
    RetVal = EnumForms(PrinterHandle, 1, temp(0), BytesNeeded, BytesNeeded, _
    NumForms)
    Call CopyMemory(aFI1(0), temp(0), BytesNeeded)
    For i = 0 To NumForms - 1
    With aFI1(i)
    If .Size.cx = FormSize.cx And .Size.cy = FormSize.cy Then
    ' Found the desired form
    FormName = PtrCtoVbString(.pName)
    FormIndex = i + 1
    Exit For
    End If
    End With
    Next i
    GetFormName = FormIndex
End Function

Public Function AddNewForm(PrinterHandle As Long, FormSize As SIZEL, _
                           FormName As String) As String
    Dim FI1 As sFORM_INFO_1
    Dim aFI1() As Byte
    Dim RetVal As Long
    
    With FI1
    .flags = 0
    .pName = FormName
    With .Size
    .cx = FormSize.cx
    .cy = FormSize.cy
    End With
    With .ImageableArea
    .left = 0
    .top = 0
    .right = FI1.Size.cx
    .bottom = FI1.Size.cy
    End With
    End With
    ReDim aFI1(Len(FI1))
    Call CopyMemory(aFI1(0), FI1, Len(FI1))
    RetVal = AddForm(PrinterHandle, 1, aFI1(0))
    If RetVal = 0 Then
        If Err.LastDllError = 5 Then
            MsgBox "You do not have permissions to add a form to " & _
            Printer.Devicename, vbExclamation, "Access Denied!"
            Else
            MsgBox "Error: " & Err.LastDllError, "Error Adding Form"
            End If
        AddNewForm = "none"
        Else
        AddNewForm = FI1.pName
    End If
End Function

Public Function PtrCtoVbString(ByVal Add As Long) As String
    Dim sTemp As String * 512, X As Long
    X = lstrcpy(sTemp, ByVal Add)
    If (InStr(1, sTemp, Chr(0)) = 0) Then
    PtrCtoVbString = ""
    Else
    PtrCtoVbString = left(sTemp, InStr(1, sTemp, Chr(0)) - 1)
    End If
End Function



Public Function sGetNombrePC()

   Dim strNombre As String
   Dim intLength As Integer
   Dim lngLength As Long
       
   strNombre = Space$(256)
   lngLength = 255
   intLength = GetComputerName(strNombre, lngLength)
   sGetNombrePC = left(strNombre, lngLength)
   strNombre = ""
       
End Function


Public Function SettearHora(Año, Mes, Dia)
        Dim lReturn As Long
        Dim lpSystemTime As SYSTEMTIME
        lpSystemTime.wYear = Año
        lpSystemTime.wMonth = Mes
        lpSystemTime.wDay = Dia
        lpSystemTime.wHour = Format(Time, "HH") + 5
        lpSystemTime.wMinute = Mid(Format(Time, "hh:mm"), 4, 2)
        lReturn = SetSystemTime(lpSystemTime)
        SettearHora = lReturn
End Function

