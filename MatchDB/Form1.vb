Public Class MatchDB
    Public OldDBTables(), NewDBTables(), NewTablesToAdd(), strOldDbPath, strNewDbPath As String
    'Public OldDBTableCol(), NewDBTableCol(), NewTableColToAdd() As String
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.OpenFileDialog1.ShowDialog()
        If Me.OpenFileDialog1.FileName <> Nothing Then Me.DBPathOld.Text = Me.OpenFileDialog1.FileName
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.OpenFileDialog1.ShowDialog()
        If Me.OpenFileDialog1.FileName <> Nothing Then Me.DBPathNew.Text = Me.OpenFileDialog1.FileName
    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim i As Integer = 0
        Dim resp As String = ""
        If Me.CnRO.Checked = True And SendToTempFolder() Then
            Call SetPasswordFields()
            Log.Show()
            If My.Computer.FileSystem.FileExists(strOldDbPath) Then
                DoLog("Removing password from old DB.")
                resp = ""
                resp = RemovePassFromDB(strOldDbPath, Me.DBPassOld.Text)
                If resp <> "" Then MsgBox(resp)
                DoLog("Compact and repairing old DB.")
                CompactAndRepairDB(strOldDbPath)
                DoLog("Setting password to old DB.")
                resp = ""
                resp = SetPassToDB(strOldDbPath, Me.DBPassOld.Text)
                If resp <> "" Then MsgBox(resp)
            End If
            If My.Computer.FileSystem.FileExists(strNewDbPath) Then
                DoLog("Removing password from new DB.")
                resp = ""
                resp = RemovePassFromDB(strNewDbPath, Me.DbPassNew.Text)
                If resp <> "" Then MsgBox(resp)
                DoLog("Compact and repairing new DB.")
                CompactAndRepairDB(strNewDbPath)
                DoLog("Setting password to new DB.")
                resp = ""
                resp = SetPassToDB(strNewDbPath, Me.DbPassNew.Text)
                If resp <> "" Then MsgBox(resp)
            End If
            Call ReceiveFromTempFolder()
            Call KillTempDBs()
            Me.DBPassOld.Text = "" : Me.DBPassOld.Enabled = True
            Me.DbPassNew.Text = "" : Me.DbPassNew.Enabled = True
            Exit Sub
        End If

        If My.Computer.FileSystem.FileExists(strOldDbPath) And My.Computer.FileSystem.FileExists(strNewDbPath) And SendToTempFolder() Then
            If MsgBox("It is highly recommended that you backup both the databases before starting the process. If you have not backed up then you can click [No] to prevent the process from starting or click [Yes] to start the upgrade.", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                Log.Show()
                DoLog("Removing password from old DB.")
                resp = "" : resp = RemovePassFromDB(strOldDbPath, Me.DBPassOld.Text) : If resp <> "" Then MsgBox(resp)
                DoLog("Removing password from new DB.")
                resp = "" : resp = RemovePassFromDB(strNewDbPath, Me.DbPassNew.Text) : If resp <> "" Then MsgBox(resp)
                DoLog("Compact and repairing old DB.")
                CompactAndRepairDB(strOldDbPath)
                DoLog("Compact and repairing new DB.")
                CompactAndRepairDB(strNewDbPath)
                DoLog("Retreiving old DB table's list.")
                OldDBTables = GetTableList(strOldDbPath)
                DoLog("Retreiving new DB table's list.")
                NewDBTables = GetTableList(strNewDbPath)
                DoLog("Finding new tables to add.")
                NewTablesToAdd = NewEntriesOfArray(OldDBTables, NewDBTables)
                DoLog("Creating 'CREATE' SQLs.")
                i = 0
                If (NewTablesToAdd IsNot Nothing AndAlso NewTablesToAdd.Length > 0) Then
                    Do Until i = NewTablesToAdd.Length
                        Log.SQLBox.Items.Add(CreateCreateTableStatement(strNewDbPath, NewTablesToAdd(i)))
                        i = i + 1
                    Loop
                End If
                On Error GoTo EscapeDo2
                DoLog("Creating 'ALTER' SQLs.")
                resp = "" : resp = ProcessColumns(strOldDbPath, strNewDbPath) : If resp <> "" Then MsgBox(resp)
EscapeDo2:
                If Log.SQLBox.Items.Count = 0 Then GoTo SetPasswordscompactandEnd
                If ShakeTheOperator("The program is going to start executing every SQL from SQL Box. Please confirm the process by") = True Then
                    DoLog("Running SQLs.")
                    Do Until Log.SQLBox.Items.Count = 0
                        DoLog("Running: " & Log.SQLBox.Items.Item(0).ToString())
                        resp = "" : resp = SQLExec(strOldDbPath, Log.SQLBox.Items.Item(0).ToString) : If resp <> "" Then MsgBox(resp)
                        Log.SQLBox.Items.RemoveAt(0)
                    Loop
                Else
                    Me.SaveFileDialog1.ShowDialog()
                    Dim saveas As String = Me.SaveFileDialog1.FileName.ToString
                    If My.Computer.FileSystem.FileExists(saveas) Then Kill(saveas)
                    i = 0
                    'MsgBox(Log.SQLBox.Items.Count.ToString)
                    Do Until i = Log.SQLBox.Items.Count.ToString
                        Dim filewrite As IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(saveas, True, System.Text.Encoding.UTF8)
                        'MsgBox(Log.SQLBox.Items.Item(i).ToString)
                        filewrite.WriteLine(Log.SQLBox.Items.Item(i).ToString)
                        filewrite.Close()
                        i = i + 1
                    Loop
                End If


SetPasswordsCompactandEnd:

                DoLog("Compact and repairing old DB.")
                CompactAndRepairDB(strOldDbPath)
                DoLog("Compact and repairing new DB.")
                CompactAndRepairDB(strNewDbPath)
                DoLog("Setting password to old DB.")
                resp = "" : resp = SetPassToDB(strOldDbPath, Me.DBPassOld.Text) : If resp <> "" Then MsgBox(resp)
                DoLog("Setting password to new DB.")
                resp = "" : resp = SetPassToDB(strNewDbPath, Me.DbPassNew.Text) : If resp <> "" Then MsgBox(resp)
                Call ReceiveFromTempFolder()
                Call KillTempDBs()
            End If
        Else
            If Not My.Computer.FileSystem.FileExists(strOldDbPath) Then MsgBox("You must enter old database location.")
            If Not My.Computer.FileSystem.FileExists(strNewDbPath) Then MsgBox("You must enter new database location.")
        End If
    End Sub
    Function CompactAndRepairDB(ByVal DBPath As String) As String
        On Error GoTo EndErr
        Dim cnn As New Microsoft.Office.Interop.Access.Dao.DBEngine()
        cnn.CompactDatabase(DBPath, DBPath & "BACKUP")
        Kill(DBPath)
        Rename(DBPath & "BACKUP", DBPath)
        Exit Function
EndErr:
        MsgBox("CAR: " & Err.Description)
    End Function
    Function ArrayHas(ByVal StringArray() As String, ByVal StringToSearch As String) As Boolean
        Dim i As Integer
        Dim isfound As Boolean = False
        Do Until i = StringArray.Length
            If StringArray(i) = StringToSearch Then
                isfound = True
                Exit Do
            End If
            i = i + 1
        Loop
        Return isfound
    End Function
    Function CreateAlterTableStatement(ByVal DBPath As String, ByVal TableName As String, ByVal ColumnName() As String) As String
        On Error GoTo EndErr
        Dim cnn As New ADODB.Connection
        Dim ColumnsSchema, PrimaryKeysSchema As ADODB.Recordset
        Dim tempsql, PrimaryKeyColumn, ColLen As String
        Dim i As Integer
        cnn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & DBPath & "';"
        cnn.Mode = ADODB.ConnectModeEnum.adModeShareExclusive
        cnn.Open()
        PrimaryKeysSchema = cnn.OpenSchema(ADODB.SchemaEnum.adSchemaPrimaryKeys)
        PrimaryKeysSchema.Filter = "TABLE_NAME = '" & TableName & "'"
        If PrimaryKeysSchema.EOF = False Then PrimaryKeyColumn = PrimaryKeysSchema("COLUMN_NAME").Value
        PrimaryKeysSchema.Close()
        ColumnsSchema = cnn.OpenSchema(ADODB.SchemaEnum.adSchemaColumns)
        ColumnsSchema.Filter = "TABLE_NAME = '" & TableName & "'"
        tempsql = "ALTER TABLE `" & TableName & "` ADD"
        Do While Not ColumnsSchema.EOF
            If ArrayHas(ColumnName, ColumnsSchema("COLUMN_NAME").Value) = True Then
                If ColumnsSchema("CHARACTER_MAXIMUM_LENGTH").Value.ToString = "" Or ColumnsSchema("CHARACTER_MAXIMUM_LENGTH").Value.ToString = "0" Then ColLen = "" 'Else ColLen = "(" & ColumnsSchema("CHARACTER_MAXIMUM_LENGTH").Value & ")"
                tempsql = tempsql & " `" & ColumnsSchema("COLUMN_NAME").Value & "` " & DataCodeToName(ColumnsSchema("DATA_TYPE").Value, ColumnsSchema("CHARACTER_MAXIMUM_LENGTH").Value.ToString) & " " & ColLen ' & ColumnsSchema("IS_NULLABLE").Value & ColumnsSchema("COLUMN_DEFAULT").Value & ", " & ColumnsSchema("IS_NULLABLE").Value & ", " & DataCodeToName(ColumnsSchema("DATA_TYPE").Value) & ", " & ColumnsSchema("CHARACTER_MAXIMUM_LENGTH").Value
                If PrimaryKeyColumn = ColumnsSchema("COLUMN_NAME").Value Then
                    tempsql = tempsql + " NOT NULL IDENTITY PRIMARY KEY, "
                    'Call RemovePrimaryKeyFromTable(TableName)
                Else
                    tempsql = tempsql + ", "
                End If
            End If
            ColumnsSchema.MoveNext()
        Loop
        tempsql = tempsql.Substring(0, Len(tempsql) - 2) + ";"
        cnn.Close()
        Return tempsql
        Exit Function
EndErr:
        If cnn.State <> 0 Then cnn.Close()
        MsgBox("CATS: " & Err.Description)
    End Function
    Function RemovePrimaryKeyFromTable(ByVal TableName As String)
        Log.SQLBox.Items.Add("ALTER TABLE `" & TableName & "` DROP INDEX *;")
    End Function
    Function ProcessColumns(ByVal OLDDB As String, ByVal NEWDB As String) As String
        On Error GoTo EndErr
        OldDBTables = GetTableList(OLDDB)
        NewDBTables = GetTableList(NEWDB)

        Dim t, c As Integer
        Dim tempcolold(), tempcolnew(), coltoadd() As String

        Do Until t = OldDBTables.Length
            tempcolold = GetColumnList(OLDDB, OldDBTables(t))
            tempcolnew = GetColumnList(NEWDB, OldDBTables(t))
            coltoadd = NewEntriesOfArray(tempcolold, tempcolnew)
            If (coltoadd IsNot Nothing AndAlso coltoadd.Length > 0) Then Log.SQLBox.Items.Add(CreateAlterTableStatement(NEWDB, OldDBTables(t), coltoadd))
            t = t + 1
        Loop

        Return ""
        On Error GoTo -1
        Exit Function
EndErr:
        Return ("PC: " & Err.Description)
    End Function
    Function GetColumnList(ByVal DbPath As String, ByVal TableName As String) As String()
        On Error GoTo EndErr
        Dim cnn As New ADODB.Connection
        Dim ColumnsSchema As ADODB.Recordset
        Dim tnames(), temptnames As String
        Dim i As Integer = 0
        cnn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & DbPath & "';"
        cnn.Mode = ADODB.ConnectModeEnum.adModeShareExclusive
        cnn.Open()
        ColumnsSchema = cnn.OpenSchema(ADODB.SchemaEnum.adSchemaColumns)
        ColumnsSchema.Filter = "TABLE_NAME = '" & TableName & "'"
        Do While Not ColumnsSchema.EOF
            temptnames = temptnames & ColumnsSchema("COLUMN_NAME").Value & ","
            ColumnsSchema.MoveNext()
            i = i + 1
        Loop
        cnn.Close()
        If temptnames <> Nothing Then tnames = Split(temptnames.Substring(0, Len(temptnames) - 1), ",")
        Return tnames
        Exit Function
EndErr:
        If cnn.State <> 0 Then cnn.Close()
        MsgBox("GCL: " & Err.Description)
    End Function
    Function SQLExec(ByVal DBPath As String, ByVal SQLStatement As String) As String
        On Error GoTo EndErr
        Dim cnn As New ADODB.Connection
        cnn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & DBPath & "';"
        cnn.Mode = ADODB.ConnectModeEnum.adModeShareExclusive
        cnn.Open()
        cnn.Execute(SQLStatement)
        cnn.Close()
        Exit Function
EndErr:
        If cnn.State <> 0 Then cnn.Close()
        Return "SQLE: " & Err.Description
    End Function
    Function ShakeTheOperator(ByVal Question As String) As Boolean
        Dim authcode As String = Format(Now(), "yyHHmmMMddss")
        Dim resp As String
retry:
        resp = InputBox(Question & vbCrLf & "Enter the code below:" & authcode)
        If resp = authcode Then
            Return True
        ElseIf resp = "" Then
            Return False
        Else
            GoTo retry
        End If
    End Function
    Function CreateCreateTableStatement(ByVal DBPath As String, ByVal TableName As String) As String
        'CREATE [TEMPORARY] TABLE table (field1 type [(size)] [NOT NULL] [WITH COMPRESSION | WITH COMP] [index1] [, field2 type [(size)] [NOT NULL] [index2]
        ' [, …]] [, CONSTRAINT multifieldindex [, …]])
        On Error GoTo EndErr
        Dim cnn As New ADODB.Connection
        Dim TablesSchema, ColumnsSchema, PrimaryKeysSchema As ADODB.Recordset
        Dim tempsql, PrimaryKeyColumn, ColLen As String
        Dim i As Integer
        cnn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & DBPath & "';"
        cnn.Mode = ADODB.ConnectModeEnum.adModeShareExclusive
        cnn.Open()
        TablesSchema = cnn.OpenSchema(ADODB.SchemaEnum.adSchemaTables)
        TablesSchema.Filter = "TABLE_NAME = '" & TableName & "'"
        PrimaryKeysSchema = cnn.OpenSchema(ADODB.SchemaEnum.adSchemaPrimaryKeys)
        PrimaryKeysSchema.Filter = "TABLE_NAME = '" & TableName & "'"
        If PrimaryKeysSchema.EOF = False Then PrimaryKeyColumn = PrimaryKeysSchema("COLUMN_NAME").Value
        PrimaryKeysSchema.Close()
        ColumnsSchema = cnn.OpenSchema(ADODB.SchemaEnum.adSchemaColumns)
        ColumnsSchema.Filter = "TABLE_NAME = '" & TableName & "'"
        tempsql = "CREATE TABLE `" & TableName & "` ("
        Do While Not ColumnsSchema.EOF
            If ColumnsSchema("CHARACTER_MAXIMUM_LENGTH").Value.ToString = "" Or ColumnsSchema("CHARACTER_MAXIMUM_LENGTH").Value.ToString = "0" Then ColLen = "" 'Else ColLen = "(" & ColumnsSchema("CHARACTER_MAXIMUM_LENGTH").Value & ")"
            tempsql = tempsql & "`" & ColumnsSchema("COLUMN_NAME").Value & "` " & DataCodeToName(ColumnsSchema("DATA_TYPE").Value, ColumnsSchema("CHARACTER_MAXIMUM_LENGTH").Value.ToString) & " " & ColLen ' & ColumnsSchema("IS_NULLABLE").Value & ColumnsSchema("COLUMN_DEFAULT").Value & ", " & ColumnsSchema("IS_NULLABLE").Value & ", " & DataCodeToName(ColumnsSchema("DATA_TYPE").Value) & ", " & ColumnsSchema("CHARACTER_MAXIMUM_LENGTH").Value
            If PrimaryKeyColumn = ColumnsSchema("COLUMN_NAME").Value Then tempsql = tempsql + " NOT NULL IDENTITY PRIMARY KEY, " Else tempsql = tempsql + ", "
            ColumnsSchema.MoveNext()
        Loop
        tempsql = tempsql.Substring(0, Len(tempsql) - 2) + ");"
        cnn.Close()
        Return tempsql
        Exit Function
EndErr:
        If cnn.State <> 0 Then cnn.Close()
        MsgBox("CCTS: " & Err.Description)
    End Function
    Function DataCodeToName(ByVal DataCode As Integer, Optional ByVal ColumnLength As String = "") As String
        Select Case DataCode
            Case 205
                Return "LONGBINARY"
            Case 128
                Return "OLEOBJECT"
            Case 11
                Return "BIT"
                'Return "COUNTER"
            Case 6
                Return "CURRENCY"
            Case 135
                Return "DateTime"
            Case 7
                Return "DateTime"
            Case 72
                Return "GUID"
            Case 201 Or 202
                Return "LONGTEXT"
            Case 4
                Return "SINGLE"
            Case 5
                Return "DOUBLE"
            Case 17
                Return "BYTE"
            Case 2
                Return "SHORT"
            Case 3
                Return "LONG"
            Case 131
                Return "DECIMAL"
            Case 200
                Return "VARCHAR"
            Case 201
                Return "VARCHAR"
            Case 203
                Return "VARCHAR"
            Case 130
                Select Case ColumnLength
                    Case 0
                        Return "MEMO"
                    Case Else
                        Return "VARCHAR"
                End Select
            Case 204
                Return "VARBINARY"
            Case Else
                Return DataCode
        End Select
    End Function
    Function NewEntriesOfArray(ByVal SmallArray() As String, ByVal BigArray() As String) As String()
        Dim tnames(), temptnames As String
        Dim i, i2 As Integer
        Dim isfound As Boolean = False
        i = 0 : i2 = 0

        If Not (BigArray IsNot Nothing AndAlso BigArray.Length > 0) Then Exit Function
        If Not (SmallArray IsNot Nothing AndAlso SmallArray.Length > 0) Then
            If (BigArray IsNot Nothing AndAlso BigArray.Length > 0) Then
                Return BigArray
                Exit Function
            Else
                Exit Function
            End If
        End If

        Do Until i = BigArray.Length
            Do Until i2 = SmallArray.Length
                If SmallArray(i2).ToString.ToLower = BigArray(i).ToString.ToLower Then
                    isfound = True
                    Exit Do
                End If
                i2 = i2 + 1
            Loop
            If Not isfound = True Then temptnames = temptnames & BigArray(i) & ","
            isfound = False
            i2 = 0
            i = i + 1
        Loop
        If temptnames = "" Then
            Exit Function
        End If
        tnames = Split(temptnames.Substring(0, Len(temptnames) - 1), ",")
        Return tnames
        Exit Function
EndErr:
        MsgBox("NEOA: " & Err.Description)
    End Function
    Function GetTableList(ByVal DbPath As String) As String()
        On Error GoTo EndErr
        Dim cnn As New ADODB.Connection
        Dim TablesSchema As ADODB.Recordset
        Dim ColumnsSchema As ADODB.Recordset
        Dim tnames(), temptnames As String
        Dim i As Integer = 0
        cnn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & DbPath & "';"
        cnn.Mode = ADODB.ConnectModeEnum.adModeShareExclusive
        cnn.Open()
        TablesSchema = cnn.OpenSchema(ADODB.SchemaEnum.adSchemaTables)
        TablesSchema.Filter = "TABLE_TYPE = 'TABLE'"
        Do While Not TablesSchema.EOF
            temptnames = temptnames & TablesSchema("TABLE_NAME").Value & ","
            TablesSchema.MoveNext()
            i = i + 1
        Loop
        cnn.Close()
        tnames = Split(temptnames.Substring(0, Len(temptnames) - 1), ",")
        Return tnames
        Exit Function
EndErr:
        If cnn.State <> 0 Then cnn.Close()
        MsgBox("GTL: " & Err.Description)
    End Function
    Function RemovePassFromDB(ByVal DbPath As String, ByVal DbPassword As String) As String
        On Error GoTo EndErr
        Dim cnn As New ADODB.Connection
        cnn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & DbPath & "';"
        If DbPassword <> "" Then cnn.ConnectionString = cnn.ConnectionString & " Jet OLEDB:Database Password=" & DbPassword & ";"
        cnn.Mode = ADODB.ConnectModeEnum.adModeShareExclusive
        cnn.Open()
        If DbPassword = "" Then DbPassword = "NULL" Else DbPassword = "[" & DbPassword & "]"
        cnn.Execute("ALTER DATABASE PASSWORD NULL " & DbPassword & ";")
        cnn.Close()
        Exit Function
EndErr:
        If cnn.State <> 0 Then cnn.Close()
        Return "RPTD: " & Err.Description
    End Function
    Function SetPassToDB(ByVal DbPath As String, ByVal DbPassword As String)
        On Error GoTo EndErr
        Dim cnn As New ADODB.Connection
        cnn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & DbPath & "';"
        If DbPassword <> "" Then cnn.ConnectionString = cnn.ConnectionString & " Jet OLEDB:Database Password=" & DbPassword & ";"
        cnn.Mode = ADODB.ConnectModeEnum.adModeShareExclusive
        cnn.Open()
        If DbPassword = "" Then DbPassword = "NULL" Else DbPassword = "[" & DbPassword & "]"
        cnn.Execute("ALTER DATABASE PASSWORD " & DbPassword & " NULL;")
        cnn.Close()
        Exit Function
EndErr:
        If cnn.State <> 0 Then cnn.Close()
        Return "SPTB: " & Err.Description
    End Function
    Function DoLog(ByVal Text As String)
        Log.LogBox.Text = Log.LogBox.Text & vbCrLf & Text
    End Function
    Function DoLogSQL(ByVal Text As String)
        Log.SQLBox.Text = Log.SQLBox.Text & vbCrLf & Text
    End Function
    Function SetPasswordFields()
        If Me.HasPassDbOld.Checked Then
            Me.DBPassOld.Text = ReturnPassOfDB(strOldDbPath)
            Me.DBPassOld.Enabled = False
        End If
        If Me.HasPassDbNew.Checked Then
            Me.DbPassNew.Text = ReturnPassOfDB(strNewDbPath)
            Me.DbPassNew.Enabled = False
        End If
    End Function
    Function ReturnPassOfDB(ByVal DBName As String) As String
        Dim DbFileName As String = My.Computer.FileSystem.GetName(DBName)
        DbFileName = DbFileName.ToLower
        Select Case DbFileName
            Case "SampleDBName"
                Return "thisPassword"
        End Select
    End Function
    Function SendToTempFolder() As Boolean
        On Error GoTo err_occured

        strOldDbPath = My.Computer.FileSystem.SpecialDirectories.Temp & "\" & My.Computer.FileSystem.GetName(Me.DBPathOld.Text)
        strNewDbPath = My.Computer.FileSystem.SpecialDirectories.Temp & "\" & My.Computer.FileSystem.GetName(Me.DBPathNew.Text)

        FileCopy(Me.DBPathOld.Text, strOldDbPath)
        FileCopy(Me.DBPathNew.Text, strNewDbPath)
        Return True

        Exit Function
err_occured:
        MsgBox(Err.Number & ": " & Err.Description)
        Return False
    End Function
    Function ReceiveFromTempFolder() As Boolean
        On Error GoTo err_occured

        strOldDbPath = My.Computer.FileSystem.SpecialDirectories.Temp & "\" & My.Computer.FileSystem.GetName(Me.DBPathOld.Text)
        strNewDbPath = My.Computer.FileSystem.SpecialDirectories.Temp & "\" & My.Computer.FileSystem.GetName(Me.DBPathNew.Text)

        Kill(Me.DBPathOld.Text)
        FileCopy(strOldDbPath, Me.DBPathOld.Text)
        Kill(Me.DBPathNew.Text)
        FileCopy(strNewDbPath, Me.DBPathNew.Text)
        Return True

        Exit Function
err_occured:
        MsgBox(Err.Number & ": " & Err.Description)
        Return False
    End Function
    Function KillTempDBs()
        Kill(strOldDbPath)
        Kill(strNewDbPath)
    End Function

End Class
'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\myFolder\myAccessFile.accdb;
'Jet OLEDB:Database Password=MyDbPassword;
'
'SELECT Name FROM MSysObjects WHERE Type = 1 AND LvProp = 'Long Binary Data' ORDER BY Name
'
'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\mydatabase.mdb;
'Jet OLEDB:Database Password=MyDbPassword;
'
'GRANT SELECT ON MSysObjects TO Admin;
'
'Dim Conn As New ADODB.Connection
'Dim TablesSchema As ADODB.Recordset
'Dim ColumnsSchema As ADODB.Recordset'

'Open connection you want To get database objects
'  Conn.Provider = "MSDASQL"
'  Conn.Open "DSN=...;Database=...;", "UID", "PWD"

'Get all database tables.
'  Set TablesSchema = Conn.OpenSchema(adSchemaTables) 
'  Do While Not TablesSchema.EOF
'Get all table columns.
'    Set ColumnsSchema = Conn.OpenSchema(adSchemaColumns, _
'      Array(Empty, Empty, "" & TablesSchema("TABLE_NAME")))
'    Do While Not ColumnsSchema.EOF
'      Debug.Print TablesSchema("TABLE_NAME") & ", " & _
'        ColumnsSchema("COLUMN_NAME")
'      ColumnsSchema.MoveNext
'    Loop
'    TablesSchema.MoveNext
'  Loop
'
'Constant 	Value 	Description
'AdArray 	0x2000 	A flag value, always combined with another data type constant, that indicates an array of the other data type. Does not apply to ADOX.
'adBigInt 	20 	Indicates an eight-byte signed integer (DBTYPE_I8).
'adBinary 	128 	Indicates a binary value (DBTYPE_BYTES).
'adBoolean 	11 	Indicates a Boolean value (DBTYPE_BOOL).
'adBSTR 	8 	Indicates a null-terminated character string (Unicode) (DBTYPE_BSTR).
'adChapter 	136 	Indicates a four-byte chapter value that identifies rows in a child rowset (DBTYPE_HCHAPTER).
'adChar 	129 	Indicates a string value (DBTYPE_STR).
'adCurrency 	6 	Indicates a currency value (DBTYPE_CY). Currency is a fixed-point number with four digits to the right of the decimal point. It is stored in an eight-byte signed integer scaled by 10,000.
'adDate 	7 	Indicates a date value (DBTYPE_DATE). A date is stored as a double, the whole part of which is the number of days since December 30, 1899, and the fractional part of which is the fraction of a day.
'adDBDate 	133 	Indicates a date value (yyyymmdd) (DBTYPE_DBDATE).
'adDBTime 	134 	Indicates a time value (hhmmss) (DBTYPE_DBTIME).
'adDBTimeStamp 	135 	Indicates a date/time stamp (yyyymmddhhmmss plus a fraction in billionths) (DBTYPE_DBTIMESTAMP).
'adDecimal 	14 	Indicates an exact numeric value with a fixed precision and scale (DBTYPE_DECIMAL).
'adDouble 	5 	Indicates a double-precision floating-point value (DBTYPE_R8).
'adEmpty 	0 	Specifies no value (DBTYPE_EMPTY).
'adError 	10 	Indicates a 32-bit error code (DBTYPE_ERROR).
'adFileTime 	64 	Indicates a 64-bit value representing the number of 100-nanosecond intervals since January 1, 1601 (DBTYPE_FILETIME).
'adGUID 	72 	Indicates a globally unique identifier (GUID) (DBTYPE_GUID).
'adIDispatch 	9 	Indicates a pointer to an IDispatch interface on a COM object (DBTYPE_IDISPATCH).
'
'Note This data type is currently not supported by ADO. Usage may cause unpredictable results.
'adInteger 	3 	Indicates a four-byte signed integer (DBTYPE_I4).
'adIUnknown 	13 	Indicates a pointer to an IUnknown interface on a COM object (DBTYPE_IUNKNOWN).'
'
'Note This data type is currently not supported by ADO. Usage may cause unpredictable results.
'adLongVarBinary 	205 	Indicates a long binary value.
'adLongVarChar 	201 	Indicates a long string value.
'adLongVarWChar 	203 	Indicates a long null-terminated Unicode string value.
'adNumeric 	131 	Indicates an exact numeric value with a fixed precision and scale (DBTYPE_NUMERIC).
'adPropVariant 	138 	Indicates an Automation PROPVARIANT (DBTYPE_PROP_VARIANT).
'adSingle 	4 	Indicates a single-precision floating-point value (DBTYPE_R4).
'adSmallInt 	2 	Indicates a two-byte signed integer (DBTYPE_I2).
'adTinyInt 	16 	Indicates a one-byte signed integer (DBTYPE_I1).
'adUnsignedBigInt 	21 	Indicates an eight-byte unsigned integer (DBTYPE_UI8).
'adUnsignedInt 	19 	Indicates a four-byte unsigned integer (DBTYPE_UI4).
'adUnsignedSmallInt 	18 	Indicates a two-byte unsigned integer (DBTYPE_UI2).
'adUnsignedTinyInt 	17 	Indicates a one-byte unsigned integer (DBTYPE_UI1).
'adUserDefined 	132 	Indicates a user-defined variable (DBTYPE_UDT).
'adVarBinary 	204 	Indicates a binary value.
'adVarChar 	200 	Indicates a string value.
'adVariant 	12 	Indicates an Automation Variant (DBTYPE_VARIANT).'
'
'Note This data type is currently not supported by ADO. Usage may cause unpredictable results.
'adVarNumeric 	139 	Indicates a numeric value.
'adVarWChar 	202 	Indicates a null-terminated Unicode character string.
'adWChar 	130 	Indicates a null-terminated Unicode character string (DBTYPE_WSTR).
'
'
'
'
'
'
'
'
'
'
'
'
'
'