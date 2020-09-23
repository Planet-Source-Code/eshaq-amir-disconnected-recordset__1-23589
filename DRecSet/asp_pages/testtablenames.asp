<%@ Language=VBScript %>
<%

'--- establish access connection
Sub OpenDB (ByRef con, d, Dir)
	DB = d & ".mdb"
	Path = Server.MapPath(Dir) & "\"

	DSN="DRIVER={Microsoft Access Driver (*.mdb)};"
	DSN=DSN & "DBQ=" & Path & DB & ";"

	Set con = Server.CreateObject("ADODB.Connection")
	con.Open DSN
End Sub

dbName = request.form("dbName")

OpenDB con, dbName, "../myDbPath"


   Set cat = Server.CreateObject("ADOX.Catalog") 

   Set cat.ActiveConnection = con
   
	response.write "[[--DBTables--]]" & vbcrlf  	
	for i = 0 to cat.Tables.count-1
		if cat.Tables(i).type = "TABLE" then response.write cat.Tables(i).name & vbcrlf
	next

%>