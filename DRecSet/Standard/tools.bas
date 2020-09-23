Attribute VB_Name = "tools"
Option Explicit
'-------------------------------------
' DRecSet - ActiveX Control that sends database requests to an asp page on a server on the internet
' Author : Amir Eshaq [jamalag@hotmail.com or aeshaq@wideframe.com]
' Date : 29th May 2001
'
' Purpose : To return disconnected recordset based on a supplied query.
'
' Requires testtable.asp and testtablenames.asp to reside on a server in the internet.
' testtablenames.asp - accepts database name and returns names of tables in that database
' testtable.asp - accepts database name and a SQL statement to execute. returns results of the SQL
'
' The code is straight forward and quite simple to understand. It demonstrates the
' use of Winsock, Listbox, and datagrid controls.
'
'-------------------------------------
'

Function tableList(ByRef KSresults As String)
' This function will format the string received through KSresults and add as items (table names) into
' the lstTables

Dim mymainarray, myfieldarray, myTablearray, i As Integer


'mymainarray(0) - header, we are not interested in this.
'mymainarray(1) - the data that we need, contains all the names of tables on the database we have requested
mymainarray = Split(KSresults, "[[--DBTables--]]")

'myTablearray(1,2,3,4,...) - this array will hold all the table names
myTablearray = Split(mymainarray(1), vbCrLf)


'add tabel names to list box lstTables
For i = 1 To UBound(myTablearray) - 1
        UserControl1.lstTables.AddItem myTablearray(i)
        UserControl1.lstTables.ItemData(UserControl1.lstTables.NewIndex) = i - 1
Next

End Function

Function queryResults(ByRef KSresults As String)
' This function will format the string received through KSresults into
' Field Names, table records. It will add the field names into lstFields, append records into
' the recordset rs and finally set the datasource of the data grid to rs

Dim rs As ADODB.Recordset
Dim arrayval As String
Dim mymainarray, myfieldarray, myfieldname As String, myfieldnames, mydataarray, myarray, myarray2, i As Integer, X As Integer, errSubst As Integer


' in case of an error jump to tag myError
' example of error would be appending two fields with the same name.
'
On Error GoTo myError

' mymainarray(0) - header, we are not interested in this.
' mymainarray(1) - the data that we need, contains field names and records from the query sSQL
mymainarray = Split(KSresults, "[[--fieldnamestart--]]")

' myfieldarray(0) - will contain field names
' myfieldarray(1) - will contain all records as one string
myfieldarray = Split(mymainarray(1), "[[--fieldnameend--]]" & vbCrLf)

' myfieldnames - this array now has individual field names
myfieldnames = Split(myfieldarray(0), vbCrLf)

' create a new recordset to hold our disconnected recordset
Set rs = New ADODB.Recordset

' clear the list on list box lstFields
UserControl1.lstFields.Clear

' loop through the fields names array
For i = 0 To UBound(myfieldnames) - 1

    ' myfieldname holds the current field name. incase of an error while appending this field name to
    ' the recordset, then the value of i is appended to the current field name
    ' so that there would be no duplicates
    myfieldname = myfieldnames(i)
    rs.Fields.Append myfieldname, adChar, 255, adFldFixed

        ' populate the lstField with the field names
        UserControl1.lstFields.AddItem myfieldname
        UserControl1.lstFields.ItemData(UserControl1.lstFields.NewIndex) = i

Next

' open the recordset
rs.Open

' this array will hold individual record in string format
mydataarray = Split(myfieldarray(1), vbCrLf & vbCrLf & "[[--end--]]")

' loop through the unformatted records' array
For i = 0 To UBound(mydataarray) - 1

    ' split each item in the records array (mydataarray) into data values.
    myarray2 = Split(mydataarray(i), vbCrLf & "[[--fld--]]")
    
    rs.AddNew
    
    ' loop through all the cells in a record andd them to the recordset
    For X = 0 To UBound(myarray2)
    
        arrayval = myarray2(X)
        rs(X) = Trim(freplaceSpecialChar(arrayval))
        
    Next
    
Next

Set UserControl1.DataGrid2.DataSource = rs

Set rs = Nothing


myError:

    If Err.Number = 3367 Then
    
        ' This error indicates there was a duplicate fieldnames assigned to the recordset
        myfieldname = myfieldname & i
        Resume
        
    End If

End Function

Function replaceSpecialChar(ByRef specialChar As String) As String
' This function converts certain characters to what IIS server expects
    specialChar = Replace(specialChar, "%", "%25")
    specialChar = Replace(specialChar, " ", "%20")
    specialChar = Replace(specialChar, "&", "%26")
    specialChar = Replace(specialChar, vbCrLf, "%0D%0A")

replaceSpecialChar = specialChar

End Function

Function freplaceSpecialChar(ByRef specialChar As String) As String
' This function may not be necessary - I havent really tested this
    specialChar = Replace(specialChar, "%20", " ")
    specialChar = Replace(specialChar, "%26", "&")
    'specialChar = Replace(specialChar, "%0D%0A", vbCrLf)
    'specialChar = Replace(specialChar, "%25", "%")

freplaceSpecialChar = specialChar

End Function
