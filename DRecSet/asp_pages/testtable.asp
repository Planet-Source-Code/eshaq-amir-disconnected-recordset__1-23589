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
sSQL = request.form("sSQL")

OpenDB con, dbName, "../myDbPath"

set rs = con.Execute(sSQL)
Fieldcount = 0

response.write "[[--fieldnamestart--]]"
for each fld in rs.fields
	Fieldcount = Fieldcount + 1
	response.write fld.name & vbcrlf
next
response.write "[[--fieldnameend--]]" & vbcrlf


if rs.bof and rs.eof then
	response.write "no results!"
	response.write rs(Fieldcount-1) & " " & vbcrlf & vbcrlf & "[[--end--]]"
else
	do until rs.eof
		for i = 0 to Fieldcount-2
			response.write rs(i) & vbcrlf & "[[--fld--]]"
		next
		response.write rs(Fieldcount-1) & " " & vbcrlf & vbcrlf & "[[--end--]]"
	rs.movenext
	loop
end if



%>