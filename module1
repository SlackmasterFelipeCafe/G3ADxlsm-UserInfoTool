' Dump-o-matic.xlsm         Scott Brondel
' 8/29/2013                 sbrondel@usgs.gov
'
' .5 beta   8/29/2013
' - still to do - incorporate Location Name, Office Building Name, Mail Stop,
'   Room Number, Latitude, Longitude

' .6  12/2/2013
' - added in AD attribute title as Job Title, department as Sub Bureau description
'
' .7 12/5/2013
' - added in AD attribute canonicalName as User's OU Path


Sub RunMe_Click()

    Dim newSheet As Excel.Worksheet
    Dim baseSheet As Excel.Worksheet
    Dim debugSheet As Excel.Worksheet
    
    Set debugSheet = Sheets("Debug")
    Set baseSheet = Sheets(1)
    
    baseSheet.Range("A10:A20").Value = ""
    baseSheet.Range("A10").Value = "Status:"
    
    ' Turn off screen updating
    Application.ScreenUpdating = False

    ' Add a new worksheet.
    Sheets.Add After:=Sheets(Sheets.Count)
    
    ' Set newSheet variable to the new worksheet.
    Set newSheet = ActiveSheet
    strName = Date & " " & Time
    strName = Replace(strName, ":", "_")
    strName = Replace(strName, "/", "_")
    newSheet.Name = strName
    
    ' Go back to baseSheet
    baseSheet.Activate
    Application.ScreenUpdating = True
    
    ' Change format to Text
    newSheet.Columns("A:ZZ").NumberFormat = "@"
    
    ' Create headers
    baseSheet.Range("A11").Value = "Creating headers..."
    Create_Headers
            
    ' Populate Data
    baseSheet.Range("A12").Value = "Querying Active Directory..."
    Query_AD
    
    ' Autosize columns
    baseSheet.Range("A16").Value = "Formatting data..."
    newSheet.Columns("A:ZZ").AutoFit
   
    ' Sort data
    str_Sorter = debugSheet.Range("H8").Value
    strHeaders = debugSheet.Range("H5")
    strHeaders = Left(strHeaders, Len(strHeaders) - 2)
    arrHeaders = Split(strHeaders, ",")
    
    i = 0
    intSortColumn = 0
    Do While i < UBound(arrHeaders) + 1
        If str_Sorter = Trim(arrHeaders(i)) Then
            intSortColumn = i + 1
        End If
        i = i + 1
    Loop
    
    baseSheet.Range("A17").Value = "Sorting the data by " & str_Sorter & "..."
    'newSheet.Columns("A:ZZ").Sort Key1:=newSheet.Cells(2, intSortColumn), Order1:=xlAscending, Header:=xlYes 'Commented out becuase it seems I broke this >>-PJB-> 12/05/2017


    ' Cleanup
    baseSheet.Range("A10:A20").Value = ""
    newSheet.Activate
    

End Sub
Sub RunMe_G3SC_Click()
'Created 12/5/2017 for GGGSC Query to populate Intranet employee listing contact modifications made to RunMe_Click - >>-pbrown@usgs.gov->

    Dim newSheet As Excel.Worksheet
    Dim baseSheet As Excel.Worksheet
    Dim debugSheet As Excel.Worksheet
    
    
    Set debugSheet = Sheets("Debug")
    Set baseSheet = Sheets(1)
    
    baseSheet.Range("A10:A20").Value = ""
    baseSheet.Range("A10").Value = "Status:"
    
    ' Turn off screen updating
    Application.ScreenUpdating = False

    ' Add a new worksheet.
    Sheets.Add After:=Sheets(Sheets.Count)
    
    ' Set newSheet variable to the new worksheet.
    Set newSheet = ActiveSheet
    strName = Date & " " & Time
    strName = Replace(strName, ":", "_")
    strName = Replace(strName, "/", "_")
    newSheet.Name = strName
    
    ' Go back to baseSheet
    baseSheet.Activate
    Application.ScreenUpdating = True
    
    ' Change format to Text
    newSheet.Columns("A:ZZ").NumberFormat = "@"
    
    ' Create headers
    baseSheet.Range("A11").Value = "Creating headers..."
    G3SCCreate_Headers
            
    ' Populate Data
    baseSheet.Range("A12").Value = "Querying Active Directory..."
    Query_ADG3SC
    
    ' Autosize columns
    baseSheet.Range("A16").Value = "Formatting data..."
    newSheet.Columns("A:ZZ").AutoFit
   
    ' Sort data
    str_Sorter = debugSheet.Range("H8").Value
    strHeaders = debugSheet.Range("H5")
    strHeaders = Left(strHeaders, Len(strHeaders) - 2)
    arrHeaders = Split(strHeaders, ",")
    
    i = 0
    intSortColumn = 0
    Do While i < UBound(arrHeaders) + 1
        If str_Sorter = Trim(arrHeaders(i)) Then
            intSortColumn = i + 1
        End If
        i = i + 1
    Loop
    
    baseSheet.Range("A17").Value = "Sorting the data by " & str_Sorter & "..."
    'newSheet.Columns("A:ZZ").Sort Key1:=newSheet.Cells(2, intSortColumn), Order1:=xlAscending, Header:=xlYes


    ' Cleanup
    baseSheet.Range("A10:A20").Value = ""
    newSheet.Activate
    DeleteUnwantedRecords
    LabelBuilding
    
    

End Sub

Sub DeleteUnwantedRecords()
Dim i As Long
Dim strScienceCenter As String
Dim lRow As Long

'Added by Phil Brown (pbrown@usgs.gov) to delete anything but Minerals and Crustal Records 12/05/2017
    
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = lRow To 2 Step -1
        strScienceCenter = Cells(i, 10).Value
        If InStr(strScienceCenter, "Crustal") = 0 Then
            If InStr(strScienceCenter, "Minerals") = 0 Then
                Rows(i).EntireRow.Delete
            End If
        End If
    Next i
        
    
    'Sort by Science Center
End Sub
Sub LabelBuilding()
Dim i As Long
Dim lRow As Long
'Created by Phil Brown to replace building code with a cooresponding building name or number 12/05/2017
lRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To lRow
        If Cells(i, 7).Value = "KAC" Then Cells(i, 7).Value = "20"
        If Cells(i, 7).Value = "KAD" Then Cells(i, 7).Value = "Condemned"
        If Cells(i, 7).Value = "KCF" Then Cells(i, 7).Value = "95"
        If Cells(i, 7).Value = "KDG" Then Cells(i, 7).Value = "CU - Bldg. 6"
    Next i

End Sub
Sub Create_Headers()

    Dim newSheet As Excel.Worksheet
    Dim baseSheet As Excel.Worksheet
    Set baseSheet = Sheets(1)
    Set newSheet = Sheets(Sheets.Count)
    Set debugSheet = Sheets("Debug")
    
    strHeaders = debugSheet.Range("H5")
    strHeaders = Left(strHeaders, Len(strHeaders) - 2)
    arrHeaders = Split(strHeaders, ",")
    
    i = 0
    Do While i < UBound(arrHeaders) + 1
        newSheet.Cells(1, i + 1).Value = Trim(arrHeaders(i))
        i = i + 1
    Loop

End Sub
Sub G3SCCreate_Headers() 'Modifications to Create_Headers to suit the needs of the G3SC  >>-PJB-> 12/05/2017

    Dim newSheet As Excel.Worksheet
    Dim baseSheet As Excel.Worksheet
    Set baseSheet = Sheets(1)
    Set newSheet = Sheets(Sheets.Count)
    Set debugSheet = Sheets("Debug")
    
    'strHeaders = debugSheet.Range("H5")
    strHeaders = "Last Name, First Name, MI, Title, Email Address, Phone, Building, Room, Mailing Address, Science Center" 'Left(strHeaders, Len(strHeaders) - 2)
    arrHeaders = Split(strHeaders, ",")
    
    i = 0
    Do While i < UBound(arrHeaders) + 1
        newSheet.Cells(1, i + 1).Value = Trim(arrHeaders(i))
        i = i + 1
    Loop

End Sub


Sub Query_AD()

    Dim newSheet As Excel.Worksheet
    Dim baseSheet As Excel.Worksheet
    Set baseSheet = Sheets(1)
    Set newSheet = Sheets(Sheets.Count)
    Set debugSheet = Sheets("Debug")
    str_Attributes = debugSheet.Range("H2")
    str_Attributes = str_Attributes & "ADsPath"
    arr_attributes = Split(str_Attributes, ",")
    str_Filter = Sheets("Debug").Range("H11").Value

    Set oRootDSE = GetObject("LDAP://RootDSE")
    strLDAP = "DC=gs,DC=doi,DC=net"
    
    intRow = 1
    
    Const ADS_SCOPE_SUBTREE = 2
    Set objConnection = CreateObject("ADODB.Connection")
    Set objCommand = CreateObject("ADODB.Command")
    objConnection.Provider = ("ADsDSOObject")
    objConnection.Open "Active Directory Provider"
    objCommand.ActiveConnection = objConnection
    objCommand.Properties("Page Size") = 1000
    

    
    If (str_Filter = "All") Then
        'objCommand.CommandText = "SELECT " & str_attributes & " FROM 'LDAP://" & strLDAP & _
        '    "' WHERE objectCategory = 'person' AND objectClass = 'user' AND businessCategory = '1'"
        strFilter = "(&(objectCategory=person)(objectClass=user)(|(businessCategory=1)(businessCategory=20)(businessCategory=21)(businessCategory=30)(businessCategory=31)))"
    
    ElseIf (str_Filter = "Enabled") Then
        'objCommand.CommandText = "SELECT " & str_attributes & " FROM 'LDAP://" & strLDAP & _
        '    "' WHERE objectCategory = 'person' AND objectClass = 'user' AND businessCategory = '1'"
        strFilter = "(&(objectCategory=person)(objectClass=user)(|(businessCategory=1)(businessCategory=20)(businessCategory=21)(businessCategory=30)(businessCategory=31))(!(userAccountControl:1.2.840.113556.1.4.803:=2)))"
    
    ElseIf (str_Filter = "EnabledTemp") Then
        'objCommand.CommandText = "SELECT " & str_attributes & " FROM 'LDAP://" & strLDAP & _
        '    "' WHERE objectCategory = 'person' AND objectClass = 'user' AND businessCategory = '1'"
        strFilter = "(&(objectCategory=person)(objectClass=user)(|(businessCategory=1)(businessCategory=20)(businessCategory=21)(businessCategory=30)(businessCategory=31)))"
        
    End If
    
    strQuery = "<LDAP://" & strLDAP & ">;" & strFilter & ";" & str_Attributes & ";subtree"
    objCommand.CommandText = strQuery
       
    'MsgBox (objCommand.CommandText)
    objCommand.Properties("SearchScope") = ADS_SCOPE_SUBTREE
    Set objRecordSet = objCommand.Execute
    If objRecordSet.RecordCount > 0 Then
        baseSheet.Range("A13").Value = "Found " & objRecordSet.RecordCount & " users to parse."
        baseSheet.Range("A14").Value = "Dumping information to new sheet..."
        objRecordSet.MoveFirst
        
        Do Until objRecordSet.EOF
            strPath = objRecordSet.Fields("ADsPath").Value
            strLowerPath = LCase(strPath)
            'WScript.Echo strPath
            
            ' Global skips
            ' Skip accounts in the Messaging OU
            If InStr(strLowerPath, "ou=messaging,dc=gs") Then
               
            ElseIf InStr(strLowerPath, "ou=permanent,ou=disabled accounts,dc=gs") Then
                ' Skip Permanently Disabled accounts (aka terminated)
            
            ElseIf InStr(strLowerPath, "ou=reserved upns for name changes,ou=temporary,ou=disabled accounts,dc=gs") Then
                ' Skip accounts in the Reserved UPNs OU
                
            ElseIf InStr(strLowerPath, "ou=doiaccess,dc=gs") Then
                ' Skip accounts in the DOIAccess OUs
                
            ElseIf InStr(strLowerPath, "ou=dominosync -research needed,ou=disabled accounts,dc=gs") Then
                ' Skip accounts in the DominoSync OU
            
            ' Conditional skips based on filter
            ElseIf str_Filter = "EnabledTemp" And InStr(strLowerPath, "ou=permanent,ou=disabled accounts,dc=gs") Then
                ' Skip accounts in the perm disabled OU
                
            ElseIf str_Filter = "EnabledTemp" And InStr(strLowerPath, "ou=doiaccess,dc=gs") Then
                ' Skip accounts in the DOIAccess OU
    
    
            Else
                intRow = intRow + 1
                
                If (intRow Mod 1000) = 0 Then
                    baseSheet.Range("A15").Value = "Dumped " & intRow & " records..."
                
                End If
                
                
                i = 0
                Do While i < UBound(arr_attributes) + 1

                    Select Case Trim(arr_attributes(i))
                    Case "samAccountName"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("samAccountName").Value
                    Case "mail"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("mail").Value
                    Case "givenName"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("givenName").Value
                    Case "sn"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("sn").Value
                    Case "initials"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("initials").Value
                    Case "employeeID"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("employeeID").Value
                    Case "telephoneNumber"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("telephoneNumber").Value
                    Case "division"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("division").Value
                    Case "manager"
                        strManager = objRecordSet.Fields("manager").Value
                        If InStr(LCase(strManager), "cn=") Then
                            strManager = Replace(strManager, "CN=", "")
                            strManager = Left(strManager, InStr(LCase(strManager), ",ou=") - 1)
                        End If
                        newSheet.Cells(intRow, i + 1).Value = strManager
                    Case "title"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("title").Value
                    Case "personalTitle"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("personalTitle").Value
                    Case "suffix"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("suffix").Value
                    Case "employeeType"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("employeeType").Value
                    Case "extensionAttribute1"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("extensionAttribute1").Value
                    Case "extensionAttribute5"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("extensionAttribute5").Value
                    Case "extensionAttribute6"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("extensionAttribute6").Value
                    Case "personalPager"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("personalPager").Value
                    Case "department"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("department").Value
                    Case "extensionAttribute8"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("extensionAttribute8").Value
                    Case "description"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("description").Value
                    Case "pOPCharacterSet"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("pOPCharacterSet").Value
                    Case "importedFrom"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("importedFrom").Value
                    Case "extensionAttribute7"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("extensionAttribute7").Value
                    Case "facsimileTelephoneNumber"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("facsimileTelephoneNumber").Value
                    Case "mobile"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("mobile").Value
                    Case "pager"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("pager").Value
                    Case "otherTelephone"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("otherTelephone").Value
                    Case "primaryTelexNumber"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("primaryTelexNumber").Value
                    Case "assistant"
                        strassistant = objRecordSet.Fields("assistant").Value
                        If InStr(LCase(strassistant), "cn=") Then
                            strassistant = Replace(strassistant, "CN=", "")
                            strassistant = Left(strassistant, InStr(LCase(strassistant), ",ou=") - 1)
                        End If
                        newSheet.Cells(intRow, i + 1).Value = strassistant
                        
                    Case "roomNumber"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("roomNumber").Value
                       
                    Case "streetAddress"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("streetAddress").Value
                    Case "l"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("l").Value
                    Case "st"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("st").Value
                    Case "postalCode"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("postalCode").Value
                    Case "c"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("c").Value
                    Case "houseIdentifier"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("houseIdentifier").Value
                    Case "street"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("street").Value
                    Case "canonicalName"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("canonicalName").Value
                    Case "whenCreated"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("whenCreated").Value
                        
                        
                    End Select
                    i = i + 1
                Loop
            End If
            
            objRecordSet.MoveNext
        Loop
    
    Else
        ' no records found?
        
    End If



End Sub

Sub Query_ADG3SC() 'Modifications to Query_AD suiting the needs of the G3SC  >>-PJB-> 12/05/2017

    Dim newSheet As Excel.Worksheet
    Dim baseSheet As Excel.Worksheet
    Set baseSheet = Sheets(1)
    Set newSheet = Sheets(Sheets.Count)
    Set debugSheet = Sheets("Debug")
    '"Last Name, First Name, MI, Title, Email Address, Phone, Building, Room, Mail Stop, Description"--Hardcode str_Attributes to be what is needed for our purpose >>-PJB-> 12/05/2017
    str_Attributes = "sn, givenName, initials, title, mail, telephoneNumber, division, roomNumber, street, description, " 'debugSheet.Range("H2")
    str_Attributes = str_Attributes & "ADsPath"
    arr_attributes = Split(str_Attributes, ",")
    str_Filter = "Enabled" 'Sheets("Debug").Range("H11").Value

    Set oRootDSE = GetObject("LDAP://RootDSE")
    strLDAP = "DC=gs,DC=doi,DC=net"
    
    intRow = 1
    
    Const ADS_SCOPE_SUBTREE = 2
    Set objConnection = CreateObject("ADODB.Connection")
    Set objCommand = CreateObject("ADODB.Command")
    objConnection.Provider = ("ADsDSOObject")
    objConnection.Open "Active Directory Provider"
    objCommand.ActiveConnection = objConnection
    objCommand.Properties("Page Size") = 1000
    

    
    If (str_Filter = "All") Then
        'objCommand.CommandText = "SELECT " & str_attributes & " FROM 'LDAP://" & strLDAP & _
        '    "' WHERE objectCategory = 'person' AND objectClass = 'user' AND businessCategory = '1'"
        strFilter = "(&(objectCategory=person)(objectClass=user)(|(businessCategory=1)(businessCategory=20)(businessCategory=21)(businessCategory=30)(businessCategory=31)))"
    
    ElseIf (str_Filter = "Enabled") Then
        'objCommand.CommandText = "SELECT " & str_attributes & " FROM 'LDAP://" & strLDAP & _
        '    "' WHERE objectCategory = 'person' AND objectClass = 'user' AND businessCategory = '1'"
        strFilter = "(&(objectCategory=person)(objectClass=user)(|(businessCategory=1)(businessCategory=20)(businessCategory=21)(businessCategory=30)(businessCategory=31))(!(userAccountControl:1.2.840.113556.1.4.803:=2)))"
    
    ElseIf (str_Filter = "EnabledTemp") Then
        'objCommand.CommandText = "SELECT " & str_attributes & " FROM 'LDAP://" & strLDAP & _
        '    "' WHERE objectCategory = 'person' AND objectClass = 'user' AND businessCategory = '1'"
        strFilter = "(&(objectCategory=person)(objectClass=user)(|(businessCategory=1)(businessCategory=20)(businessCategory=21)(businessCategory=30)(businessCategory=31)))"
        
    End If
    
    strQuery = "<LDAP://" & strLDAP & ">;" & strFilter & ";" & str_Attributes & ";subtree"
    objCommand.CommandText = strQuery
       
    'MsgBox (objCommand.CommandText)
    objCommand.Properties("SearchScope") = ADS_SCOPE_SUBTREE
    Set objRecordSet = objCommand.Execute
    If objRecordSet.RecordCount > 0 Then
        baseSheet.Range("A13").Value = "Found " & objRecordSet.RecordCount & " users to parse."
        baseSheet.Range("A14").Value = "Dumping information to new sheet..."
        objRecordSet.MoveFirst
        
        Do Until objRecordSet.EOF
            strPath = objRecordSet.Fields("ADsPath").Value
            strLowerPath = LCase(strPath)
            'WScript.Echo strPath
            
            ' Global skips
            ' Skip accounts in the Messaging OU
            If InStr(strLowerPath, "ou=messaging,dc=gs") Then
               
            ElseIf InStr(strLowerPath, "ou=permanent,ou=disabled accounts,dc=gs") Then
                ' Skip Permanently Disabled accounts (aka terminated)
            
            ElseIf InStr(strLowerPath, "ou=reserved upns for name changes,ou=temporary,ou=disabled accounts,dc=gs") Then
                ' Skip accounts in the Reserved UPNs OU
                
            ElseIf InStr(strLowerPath, "ou=doiaccess,dc=gs") Then
                ' Skip accounts in the DOIAccess OUs
                
            ElseIf InStr(strLowerPath, "ou=dominosync -research needed,ou=disabled accounts,dc=gs") Then
                ' Skip accounts in the DominoSync OU
            
            ' Conditional skips based on filter
            ElseIf str_Filter = "EnabledTemp" And InStr(strLowerPath, "ou=permanent,ou=disabled accounts,dc=gs") Then
                ' Skip accounts in the perm disabled OU
                
            ElseIf str_Filter = "EnabledTemp" And InStr(strLowerPath, "ou=doiaccess,dc=gs") Then
                ' Skip accounts in the DOIAccess OU
    
    
            Else
                intRow = intRow + 1
                
                If (intRow Mod 1000) = 0 Then
                    baseSheet.Range("A15").Value = "Dumped " & intRow & " records..."
                
                End If
                
                
                i = 0
                Do While i < UBound(arr_attributes) + 1

                    Select Case Trim(arr_attributes(i))
                    Case "samAccountName"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("samAccountName").Value
                    Case "mail"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("mail").Value
                    Case "givenName"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("givenName").Value
                    Case "sn"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("sn").Value
                    Case "initials"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("initials").Value
                    Case "employeeID"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("employeeID").Value
                    Case "telephoneNumber"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("telephoneNumber").Value
                    Case "division"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("division").Value
                    Case "manager"
                        strManager = objRecordSet.Fields("manager").Value
                        If InStr(LCase(strManager), "cn=") Then
                            strManager = Replace(strManager, "CN=", "")
                            strManager = Left(strManager, InStr(LCase(strManager), ",ou=") - 1)
                        End If
                        newSheet.Cells(intRow, i + 1).Value = strManager
                    Case "title"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("title").Value
                    Case "personalTitle"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("personalTitle").Value
                    Case "suffix"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("suffix").Value
                    Case "employeeType"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("employeeType").Value
                    Case "extensionAttribute1"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("extensionAttribute1").Value
                    Case "extensionAttribute5"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("extensionAttribute5").Value
                    Case "extensionAttribute6"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("extensionAttribute6").Value
                    Case "personalPager"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("personalPager").Value
                    Case "department"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("department").Value
                    Case "extensionAttribute8"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("extensionAttribute8").Value
                    Case "description"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("description").Value
                    Case "pOPCharacterSet"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("pOPCharacterSet").Value
                    Case "importedFrom"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("importedFrom").Value
                    Case "extensionAttribute7"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("extensionAttribute7").Value
                    Case "facsimileTelephoneNumber"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("facsimileTelephoneNumber").Value
                    Case "mobile"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("mobile").Value
                    Case "pager"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("pager").Value
                    Case "otherTelephone"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("otherTelephone").Value
                    Case "primaryTelexNumber"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("primaryTelexNumber").Value
                    Case "assistant"
                        strassistant = objRecordSet.Fields("assistant").Value
                        If InStr(LCase(strassistant), "cn=") Then
                            strassistant = Replace(strassistant, "CN=", "")
                            strassistant = Left(strassistant, InStr(LCase(strassistant), ",ou=") - 1)
                        End If
                        newSheet.Cells(intRow, i + 1).Value = strassistant
                        
                    Case "roomNumber"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("roomNumber").Value
                       
                    Case "streetAddress"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("streetAddress").Value
                    Case "l"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("l").Value
                    Case "st"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("st").Value
                    Case "postalCode"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("postalCode").Value
                    Case "c"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("c").Value
                    Case "houseIdentifier"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("houseIdentifier").Value
                    Case "street"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("street").Value
                    Case "canonicalName"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("canonicalName").Value
                    Case "whenCreated"
                        newSheet.Cells(intRow, i + 1).Value = objRecordSet.Fields("whenCreated").Value
                        
                        
                    End Select
                    i = i + 1
                Loop
            End If
            
            objRecordSet.MoveNext
        Loop
    
    Else
        ' no records found?
        
    End If



End Sub

Function Get_AD_Path(samAccountName)

'On Error Resume Next

If samAccountName = "" Then
    Get_AD_Path = "N/A"
    Exit Function
End If

'MsgBox ("Looking for " & samAccountName)
Set oRootDSE = GetObject("LDAP://RootDSE")
strLDAP = "DC=gs,DC=doi,DC=net"

Const ADS_SCOPE_SUBTREE = 2
Set objConnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")
objConnection.Provider = ("ADsDSOObject")
objConnection.Open "Active Directory Provider"
objCommand.ActiveConnection = objConnection
objCommand.CommandText = "SELECT AdsPath, canonicalName, samAccountName FROM 'LDAP://" & strLDAP & "' WHERE samAccountName = '" & samAccountName & "'"

'MsgBox (objCommand.CommandText)
objCommand.Properties("SearchScope") = ADS_SCOPE_SUBTREE
Set objRecordSet = objCommand.Execute
If objRecordSet.RecordCount = 0 Then
    ' look for computer - add a $ to samAccountName
    Set objConnection = CreateObject("ADODB.Connection")
    Set objCommand = CreateObject("ADODB.Command")
    objConnection.Provider = ("ADsDSOObject")
    objConnection.Open "Active Directory Provider"
    objCommand.ActiveConnection = objConnection
    objCommand.CommandText = "SELECT AdsPath, canonicalName, samAccountName FROM 'LDAP://" & strLDAP & "' WHERE samAccountName = '" & samAccountName & "$'"
    'MsgBox (objCommand.CommandText)
    objCommand.Properties("SearchScope") = ADS_SCOPE_SUBTREE
    Set objRecordSet = objCommand.Execute
    If objRecordSet.RecordCount = 0 Then
        Get_AD_Path = "Name not found"
    Else
        objRecordSet.MoveFirst
        strADsPath = objRecordSet.Fields("AdsPath").Value
        Get_AD_Path = GetCanonical(strADsPath)
    End If
    
Else
        objRecordSet.MoveFirst
        strADsPath = objRecordSet.Fields("AdsPath").Value

        Get_AD_Path = GetCanonical(strADsPath)
End If



End Function


Function GetCanonical(strADsPath)

    'MsgBox ("Path = " & strADsPath)
    Set oADObject = GetObject(strADsPath)
    'MsgBox ("samAccountName = " & oADobject.samAccountName)
    oADObject.GetInfoEx Array("canonicalName"), 0
    accountPath = oADObject.Get("canonicalName")
    'accountPath = objRecordSet.Fields("canonicalName").Value(0)
    'MsgBox ("Canonical = " & accountPath)
    GetCanonical = accountPath

End Function



Function Get_CN(samAccountName)

On Error Resume Next

If samAccountName = "" Then
    Get_CN = "N/A"
    Exit Function
End If

'MsgBox ("Looking for " & samAccountName)
Set oRootDSE = GetObject("LDAP://RootDSE")
strLDAP = "DC=gs,DC=doi,DC=net"

Const ADS_SCOPE_SUBTREE = 2
Set objConnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")
objConnection.Provider = ("ADsDSOObject")
objConnection.Open "Active Directory Provider"
objCommand.ActiveConnection = objConnection
objCommand.CommandText = "SELECT AdsPath, cn, samAccountName FROM 'LDAP://" & strLDAP & "' WHERE samAccountName = '" & samAccountName & "'"

'MsgBox (objCommand.CommandText)
objCommand.Properties("SearchScope") = ADS_SCOPE_SUBTREE
Set objRecordSet = objCommand.Execute
If objRecordSet.RecordCount = 0 Then
    ' look for computer - add a $ to samAccountName
    Set objConnection = CreateObject("ADODB.Connection")
    Set objCommand = CreateObject("ADODB.Command")
    objConnection.Provider = ("ADsDSOObject")
    objConnection.Open "Active Directory Provider"
    objCommand.ActiveConnection = objConnection
    objCommand.CommandText = "SELECT AdsPath, cn, samAccountName FROM 'LDAP://" & strLDAP & "' WHERE samAccountName = '" & samAccountName & "$'"
    'MsgBox (objCommand.CommandText)
    objCommand.Properties("SearchScope") = ADS_SCOPE_SUBTREE
    Set objRecordSet = objCommand.Execute
    If objRecordSet.RecordCount = 0 Then
        Get_CN = "Name not found"
    Else
        objRecordSet.MoveFirst
        Get_CN = objRecordSet.Fields("cn").Value
    End If
    
Else
        objRecordSet.MoveFirst
        strADsPath = objRecordSet.Fields("AdsPath").Value

        Get_CN = objRecordSet.Fields("cn").Value
End If



End Function


Function Get_UPN(samAccountName)

On Error Resume Next

If samAccountName = "" Then
    Get_UPN = "N/A"
    Exit Function
End If

'MsgBox ("Looking for " & samAccountName)
Set oRootDSE = GetObject("LDAP://RootDSE")
strLDAP = "DC=gs,DC=doi,DC=net"

Const ADS_SCOPE_SUBTREE = 2
Set objConnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")
objConnection.Provider = ("ADsDSOObject")
objConnection.Open "Active Directory Provider"
objCommand.ActiveConnection = objConnection
objCommand.CommandText = "SELECT AdsPath, userPrincipalName, samAccountName FROM 'LDAP://" & strLDAP & "' WHERE samAccountName = '" & samAccountName & "'"

'MsgBox (objCommand.CommandText)
objCommand.Properties("SearchScope") = ADS_SCOPE_SUBTREE
Set objRecordSet = objCommand.Execute
If objRecordSet.RecordCount = 0 Then
    ' look for computer - add a $ to samAccountName
    Set objConnection = CreateObject("ADODB.Connection")
    Set objCommand = CreateObject("ADODB.Command")
    objConnection.Provider = ("ADsDSOObject")
    objConnection.Open "Active Directory Provider"
    objCommand.ActiveConnection = objConnection
    objCommand.CommandText = "SELECT AdsPath, userPrincipalName, samAccountName FROM 'LDAP://" & strLDAP & "' WHERE samAccountName = '" & samAccountName & "$'"
    'MsgBox (objCommand.CommandText)
    objCommand.Properties("SearchScope") = ADS_SCOPE_SUBTREE
    Set objRecordSet = objCommand.Execute
    If objRecordSet.RecordCount = 0 Then
        Get_UPN = "Name not found"
    Else
        objRecordSet.MoveFirst
        Get_UPN = objRecordSet.Fields("userPrincipalName").Value
    End If
    
Else
        objRecordSet.MoveFirst
        strADsPath = objRecordSet.Fields("AdsPath").Value

        Get_UPN = objRecordSet.Fields("userPrincipalName").Value
End If



End Function


Function Get_samAccountName(userPrincipalName)

On Error Resume Next

If userPrincipalName = "" Then
    Get_samAccountName = "N/A"
    Exit Function
End If

'MsgBox ("Looking for " & samAccountName)
Set oRootDSE = GetObject("LDAP://RootDSE")
strLDAP = "DC=gs,DC=doi,DC=net"

Const ADS_SCOPE_SUBTREE = 2
Set objConnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")
objConnection.Provider = ("ADsDSOObject")
objConnection.Open "Active Directory Provider"
objCommand.ActiveConnection = objConnection
objCommand.CommandText = "SELECT AdsPath, userPrincipalName, samAccountName FROM 'LDAP://" & strLDAP & "' WHERE userPrincipalName = '" & userPrincipalName & "'"

'MsgBox (objCommand.CommandText)
objCommand.Properties("SearchScope") = ADS_SCOPE_SUBTREE
Set objRecordSet = objCommand.Execute
If objRecordSet.RecordCount = 0 Then
    ' look for computer - add a $ to samAccountName
    Set objConnection = CreateObject("ADODB.Connection")
    Set objCommand = CreateObject("ADODB.Command")
    objConnection.Provider = ("ADsDSOObject")
    objConnection.Open "Active Directory Provider"
    objCommand.ActiveConnection = objConnection
    objCommand.CommandText = "SELECT AdsPath, userPrincipalName, samAccountName FROM 'LDAP://" & strLDAP & "' WHERE userPrincipalName = '" & userPrincipalName & "$'"
    'MsgBox (objCommand.CommandText)
    objCommand.Properties("SearchScope") = ADS_SCOPE_SUBTREE
    Set objRecordSet = objCommand.Execute
    If objRecordSet.RecordCount = 0 Then
        Get_samAccountName = "Name not found"
    Else
        objRecordSet.MoveFirst
        Get_samAccountName = objRecordSet.Fields("samAccountName").Value
    End If
    
Else
        objRecordSet.MoveFirst
        strADsPath = objRecordSet.Fields("AdsPath").Value

        Get_samAccountName = objRecordSet.Fields("samAccountName").Value
End If



End Function


Function Get_sam_from_mail(mail)

On Error Resume Next

If mail = "" Then
    Get_sam_from_mail = "N/A"
    Exit Function
End If

'MsgBox ("Looking for " & samAccountName)
Set oRootDSE = GetObject("LDAP://RootDSE")
strLDAP = "DC=gs,DC=doi,DC=net"

Const ADS_SCOPE_SUBTREE = 2
Set objConnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")
objConnection.Provider = ("ADsDSOObject")
objConnection.Open "Active Directory Provider"
objCommand.ActiveConnection = objConnection
objCommand.CommandText = "SELECT AdsPath, userPrincipalName, samAccountName FROM 'LDAP://" & strLDAP & "' WHERE mail = '" & mail & "'"

'MsgBox (objCommand.CommandText)
objCommand.Properties("SearchScope") = ADS_SCOPE_SUBTREE
Set objRecordSet = objCommand.Execute
If objRecordSet.RecordCount = 0 Then
    ' look for computer - add a $ to samAccountName
    Set objConnection = CreateObject("ADODB.Connection")
    Set objCommand = CreateObject("ADODB.Command")
    objConnection.Provider = ("ADsDSOObject")
    objConnection.Open "Active Directory Provider"
    objCommand.ActiveConnection = objConnection
    objCommand.CommandText = "SELECT AdsPath, userPrincipalName, samAccountName FROM 'LDAP://" & strLDAP & "' WHERE mail = '" & mail & "$'"
    'MsgBox (objCommand.CommandText)
    objCommand.Properties("SearchScope") = ADS_SCOPE_SUBTREE
    Set objRecordSet = objCommand.Execute
    If objRecordSet.RecordCount = 0 Then
        Get_sam_from_mail = "Name not found"
    Else
        objRecordSet.MoveFirst
        Get_sam_from_mail = objRecordSet.Fields("samAccountName").Value
    End If
    
Else
        objRecordSet.MoveFirst
        strADsPath = objRecordSet.Fields("AdsPath").Value

        Get_sam_from_mail = objRecordSet.Fields("samAccountName").Value
End If



End Function

Function Get_Email(samAccountName)

On Error Resume Next

If samAccountName = "" Then
    Get_Email = "N/A"
    Exit Function
End If

'MsgBox ("Looking for " & samAccountName)
Set oRootDSE = GetObject("LDAP://RootDSE")
strLDAP = "DC=gs,DC=doi,DC=net"

Const ADS_SCOPE_SUBTREE = 2
Set objConnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")
objConnection.Provider = ("ADsDSOObject")
objConnection.Open "Active Directory Provider"
objCommand.ActiveConnection = objConnection
objCommand.CommandText = "SELECT AdsPath, mail, samAccountName FROM 'LDAP://" & strLDAP & "' WHERE samAccountName = '" & samAccountName & "'"

'MsgBox (objCommand.CommandText)
objCommand.Properties("SearchScope") = ADS_SCOPE_SUBTREE
Set objRecordSet = objCommand.Execute
If objRecordSet.RecordCount = 0 Then
    ' look for computer - add a $ to samAccountName
    Set objConnection = CreateObject("ADODB.Connection")
    Set objCommand = CreateObject("ADODB.Command")
    objConnection.Provider = ("ADsDSOObject")
    objConnection.Open "Active Directory Provider"
    objCommand.ActiveConnection = objConnection
    objCommand.CommandText = "SELECT AdsPath, mail, samAccountName FROM 'LDAP://" & strLDAP & "' WHERE samAccountName = '" & samAccountName & "$'"
    'MsgBox (objCommand.CommandText)
    objCommand.Properties("SearchScope") = ADS_SCOPE_SUBTREE
    Set objRecordSet = objCommand.Execute
    If objRecordSet.RecordCount = 0 Then
        Get_Email = "Name not found"
    Else
        objRecordSet.MoveFirst
        Get_Email = objRecordSet.Fields("mail").Value
    End If
    
Else
        objRecordSet.MoveFirst
        strADsPath = objRecordSet.Fields("AdsPath").Value

        Get_Email = objRecordSet.Fields("mail").Value
End If

End Function


Function Get_AD_Status(userPrincipalName)


'On Error Resume Next

'MsgBox ("Looking for " & samAccountName)
Set oRootDSE = GetObject("LDAP://RootDSE")
strLDAP = "DC=gs,DC=doi,DC=net"

userPrincipalName = LCase(userPrincipalName)

Const ADS_SCOPE_SUBTREE = 2
Set objConnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")
objConnection.Provider = ("ADsDSOObject")
objConnection.Open "Active Directory Provider"
objCommand.ActiveConnection = objConnection
objCommand.CommandText = "SELECT AdsPath, canonicalName, userAccountControl, samAccountName FROM 'LDAP://" & strLDAP & "' WHERE userPrincipalName = '" & userPrincipalName & "'"

'MsgBox (objCommand.CommandText)
objCommand.Properties("SearchScope") = ADS_SCOPE_SUBTREE
Set objRecordSet = objCommand.Execute
If objRecordSet.RecordCount = 0 Then
    Get_AD_Status = "Name not found"
Else
        objRecordSet.MoveFirst
        strADsPath = objRecordSet.Fields("AdsPath").Value
        Set oADObject = GetObject(strADsPath)
        
        Get_AD_Status = GetStatus(oADObject.Get("userAccountControl"))
End If


End Function


Function GetStatus(userAccountControl)
    Const ADS_UF_ACCOUNTDISABLE = 2
    Const E_ADS_PROPERTY_NOT_FOUND = &H8000500D

    If userAccountControl And ADS_UF_ACCOUNTDISABLE Then  ' account is already disabled
        GetStatus = "Disabled"
    Else  ' Account is NOT disabled
        GetStatus = "Enabled"
    End If

End Function

Function Get_AD_LastLogon(userPrincipalName)

'On Error Resume Next

'MsgBox ("Looking for " & samAccountName)
Set oRootDSE = GetObject("LDAP://RootDSE")
strLDAP = "DC=gs,DC=doi,DC=net"

userPrincipalName = LCase(userPrincipalName)

Const ADS_SCOPE_SUBTREE = 2
Set objConnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")
objConnection.Provider = ("ADsDSOObject")
objConnection.Open "Active Directory Provider"
objCommand.ActiveConnection = objConnection
objCommand.CommandText = "SELECT AdsPath, canonicalName, userAccountControl, samAccountName FROM 'LDAP://" & strLDAP & "' WHERE userPrincipalName = '" & userPrincipalName & "'"

'MsgBox (objCommand.CommandText)
objCommand.Properties("SearchScope") = ADS_SCOPE_SUBTREE
Set objRecordSet = objCommand.Execute
If objRecordSet.RecordCount = 0 Then
    Get_AD_LastLogon = "Name not found"
Else
        objRecordSet.MoveFirst
        strADsPath = objRecordSet.Fields("AdsPath").Value
        Set oADObject = GetObject(strADsPath)
        
        Get_AD_LastLogon = GetLastLogon(oADObject)
End If


End Function


Function GetLastLogon(oADObject)
   On Error Resume Next
   Const ADS_UF_ACCOUNTDISABLE = 2
   Const E_ADS_PROPERTY_NOT_FOUND = &H8000500D
   'MsgBox (oADObject.cn)
   Set objLastLogon = oADObject.Get("lastLogonTimeStamp")
      ' Has the user ever logged in?  If so, then there will be no error
   If Err.Number <> E_ADS_PROPERTY_NOT_FOUND Then
      intLastLogonTime = objLastLogon.HighPart * 4294967296# + objLastLogon.LowPart
      intLastLogonTime = intLastLogonTime / (60 * 10000000)
      intLastLogonTime = intLastLogonTime / 1440
      GetLastLogon = intLastLogonTime + #1/1/1601# & " GMT"
    
   Else  ' Account's never been logged in
      Err.Clear
      GetLastLogon = "Never logged in"
   End If

   On Error GoTo 0

End Function


Function Get_Split_Delivery_Group(userPrincipalName)
    On Error Resume Next
    intMatch = 0

    
    Set oRootDSE = GetObject("LDAP://RootDSE")
    strLDAP = "DC=gs,DC=doi,DC=net"
    
    userPrincipalName = LCase(userPrincipalName)
    
    Const ADS_SCOPE_SUBTREE = 2
    Set objConnection = CreateObject("ADODB.Connection")
    Set objCommand = CreateObject("ADODB.Command")
    objConnection.Provider = ("ADsDSOObject")
    objConnection.Open "Active Directory Provider"
    objCommand.ActiveConnection = objConnection
    objCommand.CommandText = "SELECT AdsPath, canonicalName FROM 'LDAP://" & strLDAP & "' WHERE userPrincipalName = '" & userPrincipalName & "'"

    objCommand.Properties("SearchScope") = ADS_SCOPE_SUBTREE
    Set objRecordSet = objCommand.Execute
    If objRecordSet.RecordCount = 0 Then
        Get_Split_Delivery_Group = "Name not found"
    Else
        objRecordSet.MoveFirst
        strADsPath = objRecordSet.Fields("AdsPath").Value
        Set oADObject = GetObject(strADsPath)
        
        ' Make a collection containing groups
        Set colGroups = oADObject.Groups
        
        'Go through the groups, don't test for nested groups
        For Each objGroup In colGroups
            'MsgBox objGroup.cn
            If objGroup.cn = "GS_Migrated Users" Then
                intMatch = 1
                'MsgBox objGroup.cn
                
            End If
            
        Next
        
        If (intMatch) Then
            Get_Split_Delivery_Group = "In Split Delivery Group"
            'Get_Split_Delivery_Group = intMatch
        Else
            Get_Split_Delivery_Group = "NOT in Split Delivery"
            'Get_Split_Delivery_Group = intMatch
        End If

    End If

End Function


Function Get_LDAP_Path_for_User(userPrincipalName)
    On Error Resume Next
    intMatch = 0

    
    Set oRootDSE = GetObject("LDAP://RootDSE")
    strLDAP = "DC=gs,DC=doi,DC=net"
    
    userPrincipalName = LCase(userPrincipalName)
    
    Const ADS_SCOPE_SUBTREE = 2
    Set objConnection = CreateObject("ADODB.Connection")
    Set objCommand = CreateObject("ADODB.Command")
    objConnection.Provider = ("ADsDSOObject")
    objConnection.Open "Active Directory Provider"
    objCommand.ActiveConnection = objConnection
    objCommand.CommandText = "SELECT AdsPath, canonicalName FROM 'LDAP://" & strLDAP & "' WHERE userPrincipalName = '" & userPrincipalName & "'"

    objCommand.Properties("SearchScope") = ADS_SCOPE_SUBTREE
    Set objRecordSet = objCommand.Execute
    If objRecordSet.RecordCount = 0 Then
        'MsgBox "not found"
        strResult = "Name not found"
    Else
        'MsgBox "found someone"
        objRecordSet.MoveFirst
        strADsPath = objRecordSet.Fields("AdsPath").Value
        strResult = strADsPath
    End If
    
    Get_LDAP_Path_for_User = strResult


End Function


Function Get_COUAs_for_User(userPrincipalName)

   On Error Resume Next
   
    If InStr(userPrincipalName, "@") Then
        'MsgBox "Looking for " & userPrincipalName
        sADsPath = Get_LDAP_Path_for_User(userPrincipalName)
        Set oUser = GetObject(sADsPath)
        Get_COUAs_for_User = oUser.houseIdentifier

    Else
        'MsgBox "no one home"
        Get_COUAs_for_User = "Name not found"

    End If

End Function



Function Get_EmployeeID_From_SAM(samAccountName)

On Error Resume Next

If samAccountName = "" Then
    Get_EmployeeID_From_SAM = "N/A"
    Exit Function
End If

'MsgBox ("Looking for " & samAccountName)
Set oRootDSE = GetObject("LDAP://RootDSE")
strLDAP = "DC=gs,DC=doi,DC=net"

Const ADS_SCOPE_SUBTREE = 2
Set objConnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")
objConnection.Provider = ("ADsDSOObject")
objConnection.Open "Active Directory Provider"
objCommand.ActiveConnection = objConnection
objCommand.CommandText = "SELECT AdsPath, employeeID, samAccountName FROM 'LDAP://" & strLDAP & "' WHERE samAccountName = '" & samAccountName & "'"

'MsgBox (objCommand.CommandText)
objCommand.Properties("SearchScope") = ADS_SCOPE_SUBTREE
Set objRecordSet = objCommand.Execute
If objRecordSet.RecordCount = 0 Then
    ' look for computer - add a $ to samAccountName
    Set objConnection = CreateObject("ADODB.Connection")
    Set objCommand = CreateObject("ADODB.Command")
    objConnection.Provider = ("ADsDSOObject")
    objConnection.Open "Active Directory Provider"
    objCommand.ActiveConnection = objConnection
    objCommand.CommandText = "SELECT AdsPath, employeeID, samAccountName FROM 'LDAP://" & strLDAP & "' WHERE samAccountName = '" & samAccountName & "$'"
    'MsgBox (objCommand.CommandText)
    objCommand.Properties("SearchScope") = ADS_SCOPE_SUBTREE
    Set objRecordSet = objCommand.Execute
    If objRecordSet.RecordCount = 0 Then
        Get_EmployeeID_From_SAM = "Name not found"
    Else
        objRecordSet.MoveFirst
        Get_EmployeeID_From_SAM = objRecordSet.Fields("employeeID").Value
    End If
    
Else
        objRecordSet.MoveFirst
        strADsPath = objRecordSet.Fields("AdsPath").Value

        Get_EmployeeID_From_SAM = objRecordSet.Fields("employeeID").Value
End If



End Function





Function Get_AD_LastModified(userPrincipalName)

'On Error Resume Next

'MsgBox ("Looking for " & samAccountName)
Set oRootDSE = GetObject("LDAP://RootDSE")
strLDAP = "DC=gs,DC=doi,DC=net"

userPrincipalName = LCase(userPrincipalName)

Const ADS_SCOPE_SUBTREE = 2
Set objConnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")
objConnection.Provider = ("ADsDSOObject")
objConnection.Open "Active Directory Provider"
objCommand.ActiveConnection = objConnection
objCommand.CommandText = "SELECT AdsPath, canonicalName, samAccountName FROM 'LDAP://" & strLDAP & "' WHERE userPrincipalName = '" & userPrincipalName & "'"

'MsgBox (objCommand.CommandText)
objCommand.Properties("SearchScope") = ADS_SCOPE_SUBTREE
Set objRecordSet = objCommand.Execute
If objRecordSet.RecordCount = 0 Then
    Get_AD_LastModified = "Name not found"
Else
        objRecordSet.MoveFirst
        strADsPath = objRecordSet.Fields("AdsPath").Value
        Set oADObject = GetObject(strADsPath)
        oADObject.GetInfoEx Array("modifyTimeStamp"), 0
        Get_AD_LastModified = oADObject.modifyTimeStamp   ' need to set the column to DateTime formatting in Excel
End If


End Function


Function GetMailFromCN(cn)

    On Error Resume Next
    intMatch = 0

    
    Set oRootDSE = GetObject("LDAP://RootDSE")
    strLDAP = "DC=gs,DC=doi,DC=net"
    
        
    Const ADS_SCOPE_SUBTREE = 2
    Set objConnection = CreateObject("ADODB.Connection")
    Set objCommand = CreateObject("ADODB.Command")
    objConnection.Provider = ("ADsDSOObject")
    objConnection.Open "Active Directory Provider"
    objCommand.ActiveConnection = objConnection
    objCommand.CommandText = "SELECT AdsPath, mail, canonicalName FROM 'LDAP://" & strLDAP & "' WHERE cn = '" & cn & "'"

    objCommand.Properties("SearchScope") = ADS_SCOPE_SUBTREE
    Set objRecordSet = objCommand.Execute
    If objRecordSet.RecordCount = 0 Then
        'MsgBox "not found"
        strResult = "Name not found"
    Else
        'MsgBox "found someone"
        objRecordSet.MoveFirst
        strADsPath = objRecordSet.Fields("mail").Value
        strResult = strADsPath
    End If
    
    GetMailFromCN = strResult




End Function


Function GetSamFromDisplayName(displayName)

    On Error Resume Next
    intMatch = 0

    
    Set oRootDSE = GetObject("LDAP://RootDSE")
    strLDAP = "DC=gs,DC=doi,DC=net"
    
        
    Const ADS_SCOPE_SUBTREE = 2
    Set objConnection = CreateObject("ADODB.Connection")
    Set objCommand = CreateObject("ADODB.Command")
    objConnection.Provider = ("ADsDSOObject")
    objConnection.Open "Active Directory Provider"
    objCommand.ActiveConnection = objConnection
    objCommand.CommandText = "SELECT AdsPath, samAccountName, displayName FROM 'LDAP://" & strLDAP & "' WHERE objectCategory='person' AND objectClass='user' AND businessCategory=1 AND displayName='" & displayName & "'"

    objCommand.Properties("SearchScope") = ADS_SCOPE_SUBTREE
    'MsgBox objCommand.CommandText
    Set objRecordSet = objCommand.Execute

    If objRecordSet.RecordCount = 0 Then
        'MsgBox "not found"
        strResult = "Name not found"
    Else
        'MsgBox "found someone"
        objRecordSet.MoveFirst
        strADsPath = objRecordSet.Fields("samAccountName").Value
        strResult = strADsPath
    End If
    
    GetSamFromDisplayName = strResult


End Function



Function Get_OU_from_Mail(mail)

    On Error Resume Next
    intMatch = 0

    
    Set oRootDSE = GetObject("LDAP://RootDSE")
    strLDAP = "DC=gs,DC=doi,DC=net"
    
        
    Const ADS_SCOPE_SUBTREE = 2
    Set objConnection = CreateObject("ADODB.Connection")
    Set objCommand = CreateObject("ADODB.Command")
    objConnection.Provider = ("ADsDSOObject")
    objConnection.Open "Active Directory Provider"
    objCommand.ActiveConnection = objConnection
    objCommand.CommandText = "SELECT AdsPath, mail, canonicalName FROM 'LDAP://" & strLDAP & "' WHERE mail = '" & mail & "' AND businessCategory = 1"

    objCommand.Properties("SearchScope") = ADS_SCOPE_SUBTREE
    Set objRecordSet = objCommand.Execute
    If objRecordSet.RecordCount = 0 Then
        'MsgBox "not found"
        strResult = "Name not found"
    Else
        'MsgBox "found someone"
        objRecordSet.MoveFirst
        strADsPath = objRecordSet.Fields("AdsPath").Value
        'MsgBox (strADsPath)
        strResult = Get_OU(strADsPath)
        
    End If

    
    Get_OU_from_Mail = strResult

End Function



Function Get_OU(sADsPath)

    ' Ugly function, but.....
    
    'MsgBox (sADsPath)
    ' Remove the domain info
    strShortPath = Left(sADsPath, InStr(LCase(sADsPath), ",dc=gs,dc=doi,dc=net") - 1)
    'MsgBox (strShortPath)
    
    ' now remove the CN
    intComma = InStr(LCase(sADsPath), ",ou")
    strShortPath = Right(strShortPath, Len(strShortPath) - intComma)
    'MsgBox (strShortPath)
    
    On Error GoTo 0
    
    'strPattern1 = "LDAP://OU=[A-Z]([A-Z]|[a-z])*[A-Z]{2}-[ABDGHIMNOSW],OU=[ECW]R,DC=gs,dc=doi,dc=net"
    'strPattern2 = "LDAP://OU=[A-Z]([A-Z]|[a-z])*[A-Z]{2}-[ABDGHIMNOSW](\s|[A-Z]|[a-z]|-)*,OU=DenverCO-M,OU=CR,DC=gs,dc=doi,dc=net"
    'strPattern3 = "LDAP://OU=[A-Z]([A-Z]|[a-z])*[A-Z]{2}-[ABDGHIMNOSW](\s|[A-Z]|[a-z]|-)*,OU=LafayetteLA-M,OU=CR,DC=gs,dc=doi,dc=net"
    'strPattern4 = "LDAP://OU=[A-Z]([A-Z]|[a-z])*[A-Z]{2}-[ABDGHIMNOSW](\s|[A-Z]|[a-z]|-)*,OU=FloridaFL-S,OU=ER,DC=gs,dc=doi,dc=net"
    'strPattern5 = "LDAP://OU=[A-Z]([A-Z]|[a-z])*[A-Z]{2}-[ABDGHIMNOSW](\s|[A-Z]|[a-z]|-)*,OU=FlagstaffAZ-M,OU=WR,DC=gs,dc=doi,dc=net"
    'strPattern6 = "LDAP://OU=[A-Z]([A-Z]|[a-z])*[A-Z]{2}-[ABDGHIMNOSW](\s|[A-Z]|[a-z]|-)*,OU=MenloParkCA-M,OU=WR,DC=gs,dc=doi,dc=net"
    
    strPattern1 = "OU=[A-Z]([A-Z]|[a-z])*[A-Z]{2}-[ABDGHIMNOSW](\s|[A-Z]|[a-z]|-)*,OU=DenverCO-M,OU=CR"
    strPattern2 = "OU=[A-Z]([A-Z]|[a-z])*[A-Z]{2}-[ABDGHIMNOSW](\s|[A-Z]|[a-z]|-)*,OU=LafayetteLA-M,OU=CR"
    strPattern3 = "OU=[A-Z]([A-Z]|[a-z])*[A-Z]{2}-[ABDGHIMNOSW](\s|[A-Z]|[a-z]|-)*,OU=FloridaFL-S,OU=ER"
    strPattern4 = "OU=[A-Z]([A-Z]|[a-z])*[A-Z]{2}-[ABDGHIMNOSW](\s|[A-Z]|[a-z]|-)*,OU=FlagstaffAZ-M,OU=WR"
    strPattern5 = "OU=[A-Z]([A-Z]|[a-z])*[A-Z]{2}-[ABDGHIMNOSW](\s|[A-Z]|[a-z]|-)*,OU=MenloParkCA-M,OU=WR"
    strPattern6 = "OU=[A-Z]([A-Z]|[a-z])*[A-Z]{2}-[ABDGHIMNOSW](\s|[A-Z]|[a-z]|-)*,OU=FloridaFL-W,OU=ER"
    strPattern7 = "OU=[A-Z]([A-Z]|[a-z])*[A-Z]{2}-[ABDGHIMNOSW],OU=[ECW]R"
    strPattern8 = "OU=[A-Z]([A-Z]|[a-z])*[A-Z]{2}-[ABDGHIMNOSW] Full Admins,OU=[ECW]R"
    strPattern9 = "OU=[A-Z]([A-Z]|[a-z])*[A-Z]{2}-[ABDGHIMNOSW](\s|[A-Z]|[a-z]|-)* Full Admins,OU=DenverCO-M Full Admins,OU=CR"
    strPattern10 = "OU=[A-Z]([A-Z]|[a-z])*[A-Z]{2}-[ABDGHIMNOSW](\s|[A-Z]|[a-z]|-)* Full Admins,OU=LafayetteLA-M Full Admins,OU=CR"
    strPattern11 = "OU=[A-Z]([A-Z]|[a-z])*[A-Z]{2}-[ABDGHIMNOSW](\s|[A-Z]|[a-z]|-)* Full Admins,OU=FloridaFL-S Full Admins,OU=ER"
    strPattern12 = "OU=[A-Z]([A-Z]|[a-z])*[A-Z]{2}-[ABDGHIMNOSW](\s|[A-Z]|[a-z]|-)* Full Admins,OU=FlagstaffAZ-M Full Admins,OU=WR"
    strPattern13 = "OU=[A-Z]([A-Z]|[a-z])*[A-Z]{2}-[ABDGHIMNOSW](\s|[A-Z]|[a-z]|-)* Full Admins,OU=MenloParkCA-M Full Admins,OU=WR"
    strPattern14 = "OU=[A-Z]([A-Z]|[a-z])*[A-Z]{2}-[ABDGHIMNOSW](\s|[A-Z]|[a-z]|-)* Full Admins,OU=FloridaFL-W Full Admins,OU=ER"
    strPattern15 = "OU=([A-Z]|[a-z])*,OU=DI"
    
    strTest1 = (TestOfficeRoot(strShortPath, strPattern1))
    strTest2 = (TestOfficeRoot(strShortPath, strPattern2))
    strTest3 = (TestOfficeRoot(strShortPath, strPattern3))
    strTest4 = (TestOfficeRoot(strShortPath, strPattern4))
    strTest5 = (TestOfficeRoot(strShortPath, strPattern5))
    strTest6 = (TestOfficeRoot(strShortPath, strPattern6))
    strTest7 = (TestOfficeRoot(strShortPath, strPattern7))
    strTest8 = (TestOfficeRoot(strShortPath, strPattern8))
    strTest9 = (TestOfficeRoot(strShortPath, strPattern9))
    strTest10 = (TestOfficeRoot(strShortPath, strPattern10))
    strTest11 = (TestOfficeRoot(strShortPath, strPattern11))
    strTest12 = (TestOfficeRoot(strShortPath, strPattern12))
    strTest13 = (TestOfficeRoot(strShortPath, strPattern13))
    strTest14 = (TestOfficeRoot(strShortPath, strPattern14))
    strTest15 = (TestOfficeRoot(strShortPath, strPattern15))
    
    
    If (strTest1 <> "Name Not Found") Then
        strOffice = strTest1
        'MsgBox ("Result1: " & strOffice)
    ElseIf (strTest2 <> "Name Not Found") Then
        strOffice = strTest2
        'MsgBox ("Result2: " & strOffice)
    ElseIf (strTest3 <> "Name Not Found") Then
        strOffice = strTest3
        'MsgBox ("Result3: " & strOffice)
    ElseIf (strTest4 <> "Name Not Found") Then
        strOffice = strTest4
        'MsgBox ("Result4: " & strOffice)
    ElseIf (strTest5 <> "Name Not Found") Then
        strOffice = strTest5
        'MsgBox ("Result5: " & strOffice)
    ElseIf (strTest6 <> "Name Not Found") Then
        strOffice = strTest6
        'MsgBox ("Result6: " & strOffice)
    ElseIf (strTest7 <> "Name Not Found") Then
        strOffice = strTest7
        'MsgBox ("Result7: " & strOffice)
    ElseIf (strTest8 <> "Name Not Found") Then
        strOffice = strTest8
        'MsgBox ("Result8: " & strOffice)
    ElseIf (strTest9 <> "Name Not Found") Then
        strOffice = strTest9
        'MsgBox ("Result9: " & strOffice)
    ElseIf (strTest10 <> "Name Not Found") Then
        strOffice = strTest10
        'MsgBox ("Result10: " & strOffice)
    ElseIf (strTest11 <> "Name Not Found") Then
        strOffice = strTest11
        'MsgBox ("Result11: " & strOffice)
    ElseIf (strTest12 <> "Name Not Found") Then
        strOffice = strTest12
        'MsgBox ("Result12: " & strOffice)
    ElseIf (strTest13 <> "Name Not Found") Then
        strOffice = strTest13
        'MsgBox ("Result13: " & strOffice)
    ElseIf (strTest14 <> "Name Not Found") Then
        strOffice = strTest14
        'MsgBox ("Result14: " & strOffice)
    ElseIf (strTest15 <> "Name Not Found") Then
        strOffice = strTest15
        'MsgBox ("Result15: " & strOffice)
    Else
        strOffice = "Name Not Found"
        Get_OU = strOffice
    End If
    
    ' Now, let's reverse the order
    ' Count how many ",OU=" 's we have
    
    intOUs = GetSubstringCount(strOffice, ",OU=", True)
    If intOUs = 1 Then
        tmpArray = Split(strOffice, ",OU=")
        tmpStr1 = tmpArray(1)
        tmpStr2 = tmpArray(0)

        tmpStr1 = Replace(tmpStr1, "OU=", "")
        tmpStr2 = Replace(tmpStr2, "OU=", "")
        
        'MsgBox (tmpStr1 & ":::" & tmpStr2)
        strOffice = tmpStr1 & "\" & tmpStr2
        
    ElseIf intOUs = 2 Then
        tmpArray = Split(strOffice, ",OU=")
        tmpStr1 = tmpArray(2)
        tmpStr2 = tmpArray(1)
        tmpStr3 = tmpArray(0)
        
        tmpStr1 = Replace(tmpStr1, "OU=", "")
        tmpStr2 = Replace(tmpStr2, "OU=", "")
        tmpStr3 = Replace(tmpStr3, "OU=", "")
        
        'MsgBox (tmpStr1 & ":::" & tmpStr2 & ":::" & tmpStr3)
        strOffice = tmpStr1 & "\" & tmpStr2 & "\" & tmpStr3
    
    End If

    Get_OU = strOffice

End Function



Function TestOfficeRoot(str, patrn)
    Dim regEx, Match, Matches, strFinal, intFound
    intFound = 0
    'Set regEx = New RegExp            ' Create regular expression.
    Set regEx = CreateObject("VBScript.RegExp") ' create regular expression
    regEx.Pattern = patrn            ' Set pattern.
    regEx.IgnoreCase = False            ' Make case sensitive.
    regEx.Global = True          ' Set Global applicability
    Set Matches = regEx.Execute(str)
    For Each Match In Matches
        'wscript.echo " - Office Root Found = " & Match.Value
        intFound = 1
        If Len(Match.Value) > Len(strFinal) Then
            strFinal = Match.Value
        End If
        'MsgBox (" - Office Root Found = " & strFinal)
    Next
    
    If intFound = 1 Then
        'MsgBox ("TestOfficeRoot Output: " & strFinal)
        TestOfficeRoot = strFinal
    Else
        TestOfficeRoot = "Name Not Found"
    End If
End Function

Function GetSubstringCount(strToSearch, strToLookFor, bolCaseSensative)
  If bolCaseSensative Then
    GetSubstringCount = UBound(Split(strToSearch, strToLookFor))
  Else
    GetSubstringCount = UBound(Split(UCase(strToSearch), UCase(strToLookFor)))
  End If
End Function



