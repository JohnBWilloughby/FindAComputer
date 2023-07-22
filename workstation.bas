Attribute VB_Name = "Module1"
Global ProxyPC          As Variant
Global PCOffice         As String
Global strSearch        As String
Global ShellResp        As Variant

Global pcName           As String
Global AdObjType        As String

Global ADFilter         As String
Global OUName           As String
Global DisplayUName     As String
Global IPFromHost       As String
Global SamName          As String
Global MCount           As Variant

Global eWorkstation     As String
Global eIPAdx           As String
Global eDescription     As String


Public Sub CollectPC()

Dim i As Long
Dim itmx As ListItem


On Error GoTo ErrRoutine
        ' NY -- Desktops and Laptops
        PCOffice = "NY"
        Set LDAPQuery = GetObject("LDAP://avh.com/OU=Desktops,OU=Computers,OU=NY,OU=OfficesWin10,OU=avh,DC=avh,DC=com")
        LDAPQuery.Filter = Array(ADFilter)
        For Each ADObject In LDAPQuery
            If InStr(LCase(ADObject.Description), strSearch) Then
                pcName = Mid(ADObject.Name, 4)
                DisplayUName = ADObject.Description
                Call HostToIP
                Set itmx = Form1.ListView1.ListItems.Add(, , pcName)
                itmx.SubItems(1) = IPFromHost
                itmx.SubItems(2) = DisplayUName
                itmx.SubItems(3) = PCOffice
                Form1.Refresh
            End If
        Next
        
        Set LDAPQuery = GetObject("LDAP://avh.com/OU=Laptops,OU=Computers,OU=NY,OU=OfficesWin10,OU=avh,DC=avh,DC=com")
        LDAPQuery.Filter = Array(ADFilter)
        For Each ADObject In LDAPQuery
            If InStr(LCase(ADObject.Description), strSearch) Then
                pcName = Mid(ADObject.Name, 4)
                DisplayUName = ADObject.Description
                Call HostToIP
                Set itmx = Form1.ListView1.ListItems.Add(, , pcName)
                itmx.SubItems(1) = IPFromHost
                itmx.SubItems(2) = DisplayUName
                itmx.SubItems(3) = PCOffice
                Form1.Refresh
            End If
        Next
        
        
        
       ' CT  -- Computers and Laptops
        PCOffice = "CT"
        Set LDAPQuery = GetObject("LDAP://avh.com/OU=Desktops,OU=Computers,OU=CT,OU=OfficesWin10,OU=avh,DC=avh,DC=com")
        LDAPQuery.Filter = Array(ADFilter)
        For Each ADObject In LDAPQuery
            If InStr(LCase(ADObject.Description), strSearch) Then
                pcName = Mid(ADObject.Name, 4)
                DisplayUName = ADObject.Description
                Call HostToIP
                Set itmx = Form1.ListView1.ListItems.Add(, , pcName)
                itmx.SubItems(1) = IPFromHost
                itmx.SubItems(2) = DisplayUName
                itmx.SubItems(3) = PCOffice
                Form1.Refresh
            End If
        Next

        Set LDAPQuery = GetObject("LDAP://avh.com/OU=Desktops,OU=Computers,OU=CT,OU=OfficesWin10,OU=avh,DC=avh,DC=com")
        LDAPQuery.Filter = Array(ADFilter)
        For Each ADObject In LDAPQuery
            If InStr(LCase(ADObject.Description), strSearch) Then
                pcName = Mid(ADObject.Name, 4)
                DisplayUName = ADObject.Description
                Call HostToIP
                Set itmx = Form1.ListView1.ListItems.Add(, , pcName)
                itmx.SubItems(1) = IPFromHost
                itmx.SubItems(2) = DisplayUName
                itmx.SubItems(3) = PCOffice
                Form1.Refresh
            End If
        Next

      ' DC -- Computers and Laptops
        PCOffice = "DC"
        Set LDAPQuery = GetObject("LDAP://avh.com/OU=Desktops,OU=Computers,OU=DC,OU=OfficesWin10,OU=avh,DC=avh,DC=com")
        LDAPQuery.Filter = Array(ADFilter)
        For Each ADObject In LDAPQuery
           If InStr(LCase(ADObject.Description), strSearch) Then
               pcName = Mid(ADObject.Name, 4)
                DisplayUName = ADObject.Description
                Call HostToIP
                Set itmx = Form1.ListView1.ListItems.Add(, , pcName)
                itmx.SubItems(1) = IPFromHost
                itmx.SubItems(2) = DisplayUName
                itmx.SubItems(3) = PCOffice
                Form1.Refresh
            End If
        Next

        Set LDAPQuery = GetObject("LDAP://avh.com/OU=Desktops,OU=Computers,OU=DC,OU=OfficesWin10,OU=avh,DC=avh,DC=com")
        LDAPQuery.Filter = Array(ADFilter)
        For Each ADObject In LDAPQuery
            If InStr(LCase(ADObject.Description), strSearch) Then
                pcName = Mid(ADObject.Name, 4)
                DisplayUName = ADObject.Description
                Call HostToIP
                Set itmx = Form1.ListView1.ListItems.Add(, , pcName)
                itmx.SubItems(1) = IPFromHost
                itmx.SubItems(2) = DisplayUName
                itmx.SubItems(3) = PCOffice
                Form1.Refresh
            End If
        Next


      ' SF -- Computers and Laptops
        PCOffice = "SF"
        Set LDAPQuery = GetObject("LDAP://avh.com/OU=Desktops,OU=Computers,OU=SF,OU=OfficesWin10,OU=avh,DC=avh,DC=com")
        LDAPQuery.Filter = Array(ADFilter)
        For Each ADObject In LDAPQuery
            If InStr(LCase(ADObject.Description), strSearch) Then
                pcName = Mid(ADObject.Name, 4)
                DisplayUName = ADObject.Description
                Call HostToIP
                Set itmx = Form1.ListView1.ListItems.Add(, , pcName)
                itmx.SubItems(1) = IPFromHost
                itmx.SubItems(2) = DisplayUName
                itmx.SubItems(3) = PCOffice
                Form1.Refresh
            End If
        Next

        Set LDAPQuery = GetObject("LDAP://avh.com/OU=Desktops,OU=Computers,OU=SF,OU=OfficesWin10,OU=avh,DC=avh,DC=com")
        LDAPQuery.Filter = Array(ADFilter)
       For Each ADObject In LDAPQuery
            If InStr(LCase(ADObject.Description), strSearch) Then
                pcName = Mid(ADObject.Name, 4)
                DisplayUName = ADObject.Description
                Call HostToIP
                Set itmx = Form1.ListView1.ListItems.Add(, , pcName)
                itmx.SubItems(1) = IPFromHost
                itmx.SubItems(2) = DisplayUName
                itmx.SubItems(3) = PCOffice
                Form1.Refresh
            End If
        Next


Exit Sub

ErrRoutine:
apErNum = Err.Number
Select Case apErrNum
    Case "-2147016646"
        MsgBox "Domain  " & DomainName & " does not exists"
End Select

End Sub


Private Sub HostToIP()
Dim cResolve As clsResolve
    Set cResolve = New clsResolve
    IPFromHost = cResolve.GetIPFromHostName(pcName)
End Sub

