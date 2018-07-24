set WshShell = WScript.CreateObject("WScript.Shell")
Dim Msg, Style, Title, Response, MyString
Msg = "¿quieres ver informacion de red?"    
Style = vbOkCancel    
Title = "informacion de red"    

Response = MsgBox(Msg, Style, Title)
If Response = vbOk Then    
   strComputer = "."
On Error Resume Next
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
    Set colAdapters = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration  WHERE IPEnabled = true") 
For Each objAdapter in colAdapters

      
 Msgbox "tarjeta de red: " & objAdapter.caption & vbNewLine & "puerta de enlace predeterminada: " & objAdapter.DefaultIPGateway(i) & vbNewLine & "direccion ip: " & objAdapter.IPAddress(i) & vbNewLine & "servidor dhcp: " & objAdapter.DHCPServer & vbNewLine & "dns: " & objAdapter.DNSDomain & vbNewLine & "direccion mac: " & objAdapter.MACAddress

    next
quit= "Ok"   
Else    
    MyString = "Cancel"    
End If

