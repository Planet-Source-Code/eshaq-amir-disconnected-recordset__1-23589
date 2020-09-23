VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form UserControl1 
   BackColor       =   &H80000000&
   Caption         =   "Disconnected Recordset"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   7170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton runQuery 
      Caption         =   "submit query"
      Height          =   315
      Left            =   5190
      TabIndex        =   6
      Top             =   4695
      Width           =   1800
   End
   Begin VB.TextBox sSQL 
      Height          =   1275
      Left            =   210
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   3180
      Width           =   6810
   End
   Begin VB.TextBox dbName 
      Height          =   315
      Left            =   1020
      TabIndex        =   3
      Text            =   "northwind"
      Top             =   165
      Width           =   4515
   End
   Begin VB.ListBox lstTables 
      Height          =   2205
      Left            =   225
      TabIndex        =   2
      Top             =   750
      Width           =   3330
   End
   Begin VB.ListBox lstFields 
      Height          =   2205
      Left            =   3570
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   750
      Width           =   3420
   End
   Begin VB.CommandButton tablesList 
      Caption         =   "submit"
      Height          =   315
      Left            =   5580
      TabIndex        =   0
      Top             =   165
      Width           =   1395
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   2295
      Left            =   210
      TabIndex        =   5
      Top             =   5055
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   4048
      _Version        =   393216
      ForeColor       =   -2147483646
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   210
      Top             =   7440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   210
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "Example: SELECT * FROM Orders o, Customers c where o.customerID = c.customerID"
      Top             =   4455
      Width           =   6780
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   $"UserControl1.frx":0000
      ForeColor       =   &H80000001&
      Height          =   555
      Left            =   210
      TabIndex        =   12
      Top             =   7440
      Width           =   6795
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      Caption         =   "Database:"
      Height          =   285
      Left            =   210
      TabIndex        =   11
      Top             =   225
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Tables"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   210
      TabIndex        =   10
      Top             =   525
      Width           =   3285
   End
   Begin VB.Label lblFields 
      Alignment       =   2  'Center
      Caption         =   "Fields"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   3480
      TabIndex        =   9
      Top             =   525
      Width           =   3510
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "SQL STATEMENT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   180
      Left            =   210
      TabIndex        =   8
      Top             =   3000
      Width           =   6795
   End
   Begin VB.Label lblResults 
      Alignment       =   2  'Center
      Caption         =   "Results"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   210
      Left            =   195
      TabIndex        =   7
      Top             =   4845
      Width           =   6810
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

' stores results (as a buffer) as they come in through the winsock.
Private datareceived As String

' since the requests are made to two different asp pages
' this variable will hold the name of the asp pages as they are requested.
Private aspPage As String

' HostSite is used to store Host name, in this example "www.knowledj.com"
Private HostSite As String


Private Sub Form_Load()

    ' Set the Host to attempt connections to
    HostSite = "www.knowledj.com"
    
End Sub

Private Sub lstTables_DblClick()

    ' Double clicking an item in the table list sets as a default SELECT query on that item(table name),
    sSQL.Text = "SELECT * FROM " & lstTables.List(lstTables.ListIndex)
    lblFields.Caption = "Fields (" & lstTables.List(lstTables.ListIndex) & ")"
    
    ' call the procedure to send request to the asp page that will execute the query
    runQuery_Click
    
End Sub


Private Sub runQuery_Click()

    ' Clicking this button connects to the server for a request to execute a SQL statement.
    ' clear datagrid by setting its datasource to nothing
    Set DataGrid2.DataSource = Nothing
    DataGrid2.ClearFields
    Winsock1.Close
    
    ' to execute our query we need to send request to the asp page set below
    ' see the asp page for details
    aspPage = "/wisdom/testtable.asp"
    
    ' attempt a connection to the server
    Winsock1.Connect HostSite, 80
    
End Sub

Private Sub tablesList_Click()

    ' Clicking this button connects to the server for a request to list names of all tables in
    ' the database that we have requested
    ' clear list box lstTables and close the winsock connection to the server
    Winsock1.Close
    lstTables.Clear
    
    ' to return the list of tables in the database we need to send request to the asp page set below
    ' see the asp page for details
    aspPage = "/wisdom/testtablenames.asp"
    
    ' attempt a connection to the server
    Winsock1.Connect HostSite, 80
    
End Sub


Private Sub Winsock1_Connect()

Dim strReqst As String

    ' create the HTTP Request.
    strReqst = "POST " & aspPage & " HTTP/1.1" & vbCrLf

    strReqst = strReqst & "Host: http://" & HostSite & vbCrLf
    
    ' force the server to close the connection after response.
    strReqst = strReqst & "Connection: Close" & vbCrLf
    
    ' since we are sending as data in a form, we need to specify the size of the content.
    strReqst = strReqst & "Content-Length: " & Len("dbName=" & dbName.Text & "&sSQL=" & replaceSpecialChar(sSQL.Text)) & vbCrLf
    
    strReqst = strReqst & "Content-Type: application/x-www-form-urlencoded" & vbCrLf
    
    strReqst = strReqst & "Accept: */*" & vbCrLf
    
    ' add a blank line that indicates the end of the request.
    strReqst = strReqst & vbCrLf

    ' This last line includes key/value pairs sent along to the server. Such key/value pairs are either
    ' as querystrings (e.g blah.asp?dbName=wisdom&sSQL=SELECT * FROM users) or data in a form.
    strReqst = strReqst & "dbName=" & dbName.Text & "&sSQL=" & replaceSpecialChar(sSQL.Text) & vbCrLf
    
    
    'send the request
    Winsock1.SendData strReqst
    
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

' declare a variable that will hold data received
' mind you this might not be the best way to do it if the data very large
' you might want to break the data coming in into records and freeing this variable after
' every hundreds of records. I dont know if this makes sense.
Dim sString As String

    ' no comments on this
    Winsock1.GetData sString, bytesTotal
    
    ' keep on concatinating data as they flow in.
    datareceived = datareceived & sString
    
    ' the way I have done it if you look in testtable.asp is that at the end of every record
    ' is this string (vbCrLf & vbCrLf & "[[--end--]]")
    If InStr(1, sString, vbCrLf & vbCrLf & "[[--end--]]") Then
    
        ' if data returned contains recordsets then function queryResults is called
        ' that will split apart the data into field names and into records.
        ' see the function for more details
        Call queryResults(datareceived)
        datareceived = ""
        Winsock1.Close
        
    ElseIf InStr(1, sString, "[[--DBTables--]]") Then
        ' if the data contains the above string in the condition then it contains names of tables
        ' in our requested database. Then what happens is that we call the function tableList that will
        ' split apart the data into table names. For more details see the function.
        Call tableList(datareceived)
        datareceived = ""
        Winsock1.Close
        
    End If
    
    
End Sub




