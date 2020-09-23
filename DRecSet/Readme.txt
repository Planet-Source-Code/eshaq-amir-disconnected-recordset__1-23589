'-------------------------------------
' DRecSet - ActiveX Control that sends database requests 
' to an asp page on a server on the internet.
'
' Author : Amir Eshaq [jamalag@hotmail.com] or [aeshaq@wideframe.com]
' Date : 29th May 2001
'
' Purpose : To return disconnected recordset based on a supplied query.
'
' Requires testtable.asp and testtablenames.asp 
' to reside on a server in the internet.
' testtablenames.asp - accepts database name and 
' returns names of tables in that database
' testtable.asp - accepts database name and a SQL 
' statement to execute. returns results of the SQL
'
' The code is straight forward and quite simple to understand. It demonstrates the
' use of Winsock, Listbox, and datagrid controls.
'
'-------------------------------------
'

The code is actually straight forward. I havent really tested for bugs but I am sure there are
as i wrote it in a rush. I just thought of the idea and realized that its worth sharing it with other great developers at PSC - it may be of benefit to some. I have put up an mdb file on the internet on my site that you can the application with.

Enjoy!


Components used
-------------
- Data grid
- Winsock


Folders
-------
ActiveX - Files of the same that can be compiled to use as an ActiveX control in web pages or within standard executables.
Standard - Files of the same that can be compiled to a standard exe file.
asp_pages - two asp pages that should reside in a server somewhere on the internet. You will need to change the 
path to where your databases are.