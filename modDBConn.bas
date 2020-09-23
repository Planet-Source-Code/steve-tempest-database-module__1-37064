Attribute VB_Name = "modDBConn"
Option Explicit

'******************************************************************
'******************************************************************
'***     Module for connection to various databases             ***
'***                                                            ***
'***                 Author Information                         ***
'***              Written By Steve Tempest                      ***
'***                Date: 19th July 2002                        ***
'***              Mail: steve@cstsoft.co.uk                     ***
'***                                                            ***
'******************************************************************
'***  Please feel free to use the code and freely distribute it ***
'***  but I would appreciate it if you could leave the Author   ***
'***  Information in place.                                     ***
'******************************************************************

'******************************************************************
'***                     NOTES                                  ***
'***                                                            ***
'***  The FillListView requires three bits of information       ***
'***  These are a) the Recordset. b)Name of ListView Control    ***
'***   c)Name of form containing the ListView control           ***
'***                                                            ***
'******************************************************************

'******************************************************************
'***            Code For database cnnection details.            ***
'******************************************************************

Public ConnectionStringA As String ' Returns the required connection string
Public db As ADODB.Connection

'Set up a User Defined Type for the required connection variables.
    Public Type Conn
            Server As String ' Name of Server (for SQL Server Connection), DSN Name for ODBC
            UID As String    ' User Name(For SQL, ODBC Optional for Access 97 & 2000)
            Pass As String   ' Password (For SQL, ODBC Optional for Access 97 & 2000)
            Initdb As String ' Initial Catalog for SQL Server,ODBC Database Name for Access 97 & 2000
    End Type
'Connection details for Microsoft SQL Server 2000
Sub MKConnStrSQL(Server, UID, Pass, Optional Initdb)
    If Initdb = "" Then
        ConnectionStringA = "Provider=SQLOLEDB.1;" & _
           "Data Source=" & UCase(Server) & ";uid=" & UID & ";password=" & Pass
    Else
        ConnectionStringA = "Provider=SQLOLEDB.1;" & _
           "Data Source=" & UCase(Server) & ";Initial Catalog=" & Initdb & ";uid=" & UID & ";password=" & Pass
   End If
   'Move to the Connection Details for SQL Server
    ConnSQL
End Sub
'Connection details for ODBC
Sub MKConnStrODBC(Server, Optional Initdb, Optional UID, Optional Pass)
    If Initdb = "" Then
        ConnectionStringA = "Provider=MSDASQL.1;" & _
           "DSN=" & Server
    Else
        ConnectionStringA = "Provider=MSDASQL.1;" & _
           "DSN=" & Server & ";Initial Catalog =" & Initdb & ";uid=" & UID & ";password=" & Pass
   End If
   'Move to the Connection Details for SQL Server
   ConnODBC
End Sub
'Connection details for Microsoft Access 97
Sub MKConnStrAccess97(Initdb, Optional UID, Optional Pass)
    If UID = "" Or Pass = "" Then
        ConnectionStringA = "Provider=Microsoft.Jet.OLEDB.3.51;" & _
           "Data Source=" & Initdb
    Else
        ConnectionStringA = "Provider=Microsoft.Jet.OLEDB.3.51;" & _
           "Data Source=" & Initdb & ";uid=" & UID & ";password=" & Pass
   End If
End Sub
'Connection details for Microsoft Access 2000
Sub MKConnStrAccess2000(Initdb, Optional UID, Optional Pass)
    If UID = "" Or Pass = "" Then
        ConnectionStringA = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & Initdb
    Else
        ConnectionStringA = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & Initdb & ";uid=" & UID & ";password=" & Pass
   End If
End Sub

'Connection for Microsoft SQL Server 2000
Sub ConnSQL()
On Error Resume Next
'Dim db As ADODB.Connection
    Set db = New ADODB.Connection
    db.ConnectionString = ConnectionStringA
    db.Open
    MsgBox "Connected To: " & frmMain.txtServer.Text & " - " & frmMain.txtInitDB.Text
End Sub
'Connection for ODBC datasource
Sub ConnODBC()
    Set db = New ADODB.Connection
    db.Open (ConnectionStringA)
    MsgBox "Connected To: " & frmMain.txtServer.Text & " - " & frmMain.txtInitDB.Text
End Sub
' Error reporting section
Sub WriteError()
' db is the name of the ADODB Connection
    Dim e As ADODB.Error
        If db.Errors.Count = 1 Then
            MsgBox "Error: " & db.Errors(0).Description
        ElseIf db.Errors.Count > 1 Then
            For Each e In db.Errors
                MsgBox e.Description
            Next e
        End If
        db.Errors.Clear
End Sub

'**********************************************************************
'***                                                                ***
'***   Start of Code for the buttons on the main form (frmMain)     ***
'***                                                                ***
'**********************************************************************
Public Sub cmdConnectSQL()

    Dim Myconn As Conn ' UDT to retrieve required details for connection
    Dim rs As ADODB.Recordset
    Dim FldCount As Integer
    
        'Set up recordset for data retrieval
        Set rs = New ADODB.Recordset
        Myconn.Server = frmMain.txtServer.Text
        Myconn.UID = frmMain.txtUID.Text
        Myconn.Pass = frmMain.txtPass.Text
        Myconn.Initdb = frmMain.txtInitDB.Text
        'Pass required information to modDBConn for connection
        MKConnStrSQL Myconn.Server, Myconn.UID, Myconn.Pass, Myconn.Initdb

        frmMain.lblModStat.Caption = ConnectionStringA
        rs.Source = "select * from Products"
        rs.ActiveConnection = db
        rs.Open
        rs.MoveFirst
        'The SQL select * from Products is used simply as a test.  Remember to
        'Change he sql before use.
        FillListView db.Execute("Select * From Products"), frmMain.lvResults, frmMain
        rs.Close
        Set rs = Nothing
        Set db = Nothing
        
End Sub

Public Sub cmdConODBC()
 Dim Myconn As Conn     ' UDT to retrieve required info for connection
        Myconn.Server = frmMain.txtServer.Text  ' Name of DSN
        Myconn.UID = frmMain.txtUID.Text        ' User Name
        Myconn.Pass = frmMain.txtPass.Text      ' Password
        Myconn.Initdb = frmMain.txtInitDB.Text  ' Default Database to use
        
        'Pass Required Info to ODBC section of modDBConn
        MKConnStrODBC Myconn.Server, Myconn.Initdb, Myconn.UID, Myconn.Pass
        frmMain.lblModStat.Caption = ConnectionStringA
        'The SQL query is executed directly in the FillListView section
        FillListView db.Execute("select * from Products"), frmMain.lvResults, frmMain
        Set db = Nothing
        
End Sub

'********************************************************************
'***                                                              ***
'***    End of Code for the buttons on the main form (frmMain)    ***
'***                                                              ***
'********************************************************************

'********************************************************************
'***   This is the code for the FillListView.  The code can be    ***
'***   pasted into another module (modDataListView) and used      ***
'***   in other projects.  A very handy little module             ***
'********************************************************************

'********************************************************************
'***                     Start of modDataListView                 ***
'********************************************************************
'
'Option Explicit
'
'Private Const msngWIDTH_OF_CHECKBOX As Single = 100
'Private Const msngCOLUMN_HEADER_EXTRA_WIDTH As Single = 200
'Private Const msngMAX_COLUMN_WIDTH As Single = 10000
'Private Const msngMIN_COLUMN_WIDTH As Single = 1000
'
'Public Sub FillListView(ByVal irstReports As ADODB.Recordset, _
'                        ByRef iolvwListView As MSComctlLib.ListView, _
'                        ByVal ifrmFormReference As Form)
''***************************************************************************************************
''*
''*  PURPOSE:    To fill a listview with contents of a recordset
''*
''***************************************************************************************************
'    Dim itmReport As MSComctlLib.ListItem
'    Dim fldColumn As ADODB.Field
'    Dim asngMaxSizes() As Single
'    Dim nColumn As Integer
'    Dim sngNewWidth As Single
'    Dim sTemp As String
'
'    On Error GoTo RecordsetError
'
'    'Ensure correct scalemode for TextWidth
'    ifrmFormReference.ScaleMode = vbTwips
'
'    With iolvwListView
'
'        'Set up listview
'        .AllowColumnReorder = True
'        .Enabled = True
'        .FullRowSelect = True
'        .HideColumnHeaders = False
'        .HideSelection = False
'        '.LabelEdit = lvwManual
'        .ListItems.Clear
'        .View = lvwReport
'        .Visible = True
'
'        'If we have no columns, exit
'        If irstReports.Fields.Count <= 0 Then
'            Exit Sub
'
'        End If
'
'        'Create storage to set max column width sizes
'        ReDim asngMaxSizes(0 To irstReports.Fields.Count)
'
'        'Clear any existing columns
'        .ColumnHeaders.Clear
'
'        'Add each column
'        For Each fldColumn In irstReports.Fields
'            .ColumnHeaders.Add , , fldColumn.Name
'
'        Next fldColumn
'
'        'Add each item to the listview
'        While Not irstReports.EOF
'
'            'Add item
'            Set itmReport = .ListItems.Add
'
'            'Set text, ensuring nulls are replaced with blanks
'            itmReport.Text = ReplaceNull(irstReports.Fields(0).Value, "")
'
'            'If the width of the text supplied is greater than the one stored,
'            'then store the larger one
'            If ifrmFormReference.TextWidth(itmReport.Text) > asngMaxSizes(0) Then
'                asngMaxSizes(0) = ifrmFormReference.TextWidth(Replace(itmReport.Text, vbCrLf, " "))
'
'            End If
'
'            'Add all remaining columns
'            For nColumn = 1 To irstReports.Fields.Count - 1
'
'                'Set text
'                itmReport.SubItems(nColumn) = ReplaceNull(irstReports.Fields(nColumn).Value, "")
'
'                'Generate string
'                sTemp = Replace(itmReport.SubItems(nColumn), vbCrLf, " ")
'
'                'If the column is a LongVarChar, just use the 1st 50 chars
'                If irstReports.Fields(nColumn).Type = adLongVarChar Then
'                    sTemp = Left(sTemp, 50)
'                End If
'
'                'If the width of the text supplied is greater than the one stored,
'                'then store the larger one
'                If ifrmFormReference.TextWidth(sTemp) > asngMaxSizes(nColumn) Then
'                    asngMaxSizes(nColumn) = ifrmFormReference.TextWidth(sTemp)
'
'                End If
'
'            Next nColumn
'
'            'Next item
'            irstReports.MoveNext
'
'        Wend
'
'        'Loop through each column ensuring it's size is correct
'        For nColumn = 1 To .ColumnHeaders.Count
'
'            'Check that the width of the column header is not wider than the text of it's subitems
'            If ifrmFormReference.TextWidth(.ColumnHeaders(nColumn).Text) > asngMaxSizes(nColumn - 1) Then
'                'Set width to be width of column header text
'                sngNewWidth = ifrmFormReference.TextWidth(.ColumnHeaders(nColumn).Text)
'
'            Else
'                sngNewWidth = asngMaxSizes(nColumn - 1)
'
'            End If
'
'            'Add width to account for column header borders
'            sngNewWidth = sngNewWidth + msngCOLUMN_HEADER_EXTRA_WIDTH
'
'            'Ensure column width does not get too big...
'            If sngNewWidth > msngMAX_COLUMN_WIDTH Then
'                sngNewWidth = msngMAX_COLUMN_WIDTH
'            End If
'
'            '... or small
'            If sngNewWidth <= 0 Then
'                sngNewWidth = msngMIN_COLUMN_WIDTH
'            End If
'
'            'Check first column for checkboxes, if they are used widen column
'            If nColumn = 1 Then
'                .ColumnHeaders(nColumn).Width = sngNewWidth + msngWIDTH_OF_CHECKBOX
'
'            Else
'                .ColumnHeaders(nColumn).Width = sngNewWidth
'
'            End If
'
'        Next nColumn
'
'    End With
'
'    Exit Sub
'
'RecordsetError:
'    MsgBox "FLV Error while downloading data: " & Err.Description, vbOKOnly + vbCritical, "Error"
'
'End Sub
'
'Private Function ReplaceNull(ByVal ivntValue As Variant, _
'                             ByVal isDefaultValue As String) As String
''***************************************************************************************************
''*
''*  PURPOSE:    To replace a null value with a replaceable string or return the actual string
''*
''***************************************************************************************************
'    On Error GoTo IsNull
'
'    ReplaceNull = CStr(ivntValue)
'
'    Exit Function
'
'IsNull:
'    ReplaceNull = isDefaultValue
'
'End Function

'*****************************************************************
'***                      End of modDataListView               ***
'*****************************************************************

