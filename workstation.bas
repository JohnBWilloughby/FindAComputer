Attribute VB_Name = "Module1"
Global usr_id           As String           ' the user id for the database
Global pass             As String           ' the password if used in your database
Global mySqlIP          As String           'the ip address of the machine with the mySql
Global strDataBaseName  As String

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
        ' New York -- Workstatoins and Laptops
        PCOffice = "NY"
        Set LDAPQuery = GetObject("LDAP://herzfeld-rubin.com.int/OU=workstations,OU=New York,DC=herzfeld-rubin,DC=com,DC=int")
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
        
        Set LDAPQuery = GetObject("LDAP://herzfeld-rubin.com.int/OU=Laptops,OU=New York,DC=herzfeld-rubin,DC=com,DC=int")
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
        
        ' New Jersey -- Workstatoins and Laptops
        PCOffice = "NJ"
        Set LDAPQuery = GetObject("LDAP://herzfeld-rubin.com.int/OU=workstations,OU=New Jersey,DC=herzfeld-rubin,DC=com,DC=int")
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
     
        Set LDAPQuery = GetObject("LDAP://herzfeld-rubin.com.int/OU=Laptops,OU=New Jersey,DC=herzfeld-rubin,DC=com,DC=int")
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
        
        ' Long Island -- Workstatoins and Laptops
        PCOffice = "LI"
        Set LDAPQuery = GetObject("LDAP://herzfeld-rubin.com.int/OU=workstations,OU=Long Island,DC=herzfeld-rubin,DC=com,DC=int")
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
                
        Set LDAPQuery = GetObject("LDAP://herzfeld-rubin.com.int/OU=Laptops,OU=Long Island,DC=herzfeld-rubin,DC=com,DC=int")
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

