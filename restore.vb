Imports System.Data.Odbc
Imports System.Text.RegularExpressions
Imports System
Imports System.IO

Public Class frmRestore
    Public Structure fieldStruct
        Public fid As String
        Public label As String
        Public type As String
        Public parentFieldID As String
        Public unique As Boolean
        Public required As Boolean
        Public base_type As String
        Public decimal_places As Integer
    End Structure
    Private Const AppName = "QuNectRestore"
    Private fieldNodes As Dictionary(Of String, fieldStruct)
    Private clist As String()
    Private fieldTypes As String()
    Private rids As HashSet(Of String)
    Private keyfid As String
    Private uniqueExistingFieldValues As Dictionary(Of String, HashSet(Of String))
    Private uniqueNewFieldValues As Dictionary(Of String, HashSet(Of String))
    Private sourceLabelToFieldType As Dictionary(Of String, String)
    Private sourceFieldNames As New Dictionary(Of String, Integer)
    Private destinationLabelsToFids As Dictionary(Of String, String)
    Private isBooleanTrue As Regex = New Regex("y|tr|c|[1-9]", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
    Private Const fieldNameToContainOldRecordIdNumbers = "QuNect Restore Temporary Record ID#"
    Private Class qdbVersion
        Public year As Integer
        Public major As Integer
        Public minor As Integer
    End Class
    Private qdbVer As qdbVersion = New qdbVersion
    Enum mapping
        source
        destination
    End Enum
    Enum filter
        source
        booleanOperator
        criteria
    End Enum
    Enum comparisonResult
        equal
        notEqual
        greater
        less
        null
    End Enum
    Enum errorTabs
        missing
        conversions
        required
        unique
        malformed
    End Enum
    Enum BulkOrSingle
        Neither = 0
        One = 1
        TwoOrMore = 2
    End Enum
    Enum PasswordOrToken
        Neither = 0
        password = 1
        token = 2
    End Enum
    Private cmdLineArgs() As String

    Private Sub restore_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim myBuildInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(Application.ExecutablePath)
        Me.Text = "QuNect Restore " & myBuildInfo.ProductVersion
        txtUsername.Text = GetSetting(AppName, "Credentials", "username")
        cmbPassword.SelectedIndex = CInt(GetSetting(AppName, "Credentials", "passwordOrToken", CStr(PasswordOrToken.Neither)))
        txtPassword.Text = GetSetting(AppName, "Credentials", "password")
        txtServer.Text = GetSetting(AppName, "Credentials", "server", "")
        txtAppToken.Text = GetSetting(AppName, "Credentials", "apptoken", "")
        cmbBulkorSingle.SelectedIndex = CInt(GetSetting(AppName, "config", "bulkOrSingle", CStr(BulkOrSingle.Neither)))
        lblFile.Text = GetSetting(AppName, "config", "file", "")
        lblTable.Text = GetSetting(AppName, "config", "table", "")
        Dim detectProxySetting As String = GetSetting(AppName, "Credentials", "detectproxysettings", "0")
        If detectProxySetting = "1" Then
            ckbDetectProxy.Checked = True
        Else
            ckbDetectProxy.Checked = False
        End If
        Dim firstRowHasHeaders As String = GetSetting(AppName, "config", "firstrowhasheaders", "1")
        If firstRowHasHeaders = "1" Then
            chkBxHeaders.Checked = True
        Else
            chkBxHeaders.Checked = False
        End If

        Dim ops As String() = New String() {"has any value", "equals", "does not equal", "greater than", "greater than or equal", "less than", "less than or equal", "contains", "does not contain", "starts with", "does not start with", "is null", "is not null"}
        For i As Integer = 0 To ops.Count - 1
            DirectCast(dgCriteria.Columns(filter.booleanOperator), System.Windows.Forms.DataGridViewComboBoxColumn).Items.Add(ops(i))
        Next
        dgMapping.Anchor = AnchorStyles.Bottom And AnchorStyles.Right
        dgMapping.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        'dgMapping.Dock = DockStyle.Fill
        cmdLineArgs = System.Environment.GetCommandLineArgs()
        If cmdLineArgs.Length = 5 Then
            lblFile.Text = cmdLineArgs(1)
            txtUsername.Text = cmdLineArgs(2)
            txtPassword.Text = cmdLineArgs(3)
            txtServer.Text = cmdLineArgs(4)
            cmbPassword.SelectedIndex = PasswordOrToken.token
            cmbBulkorSingle.SelectedIndex = BulkOrSingle.TwoOrMore
            Import()
            Me.Close()
        End If
    End Sub
    Private Function getConnectionString(usefids As Boolean, allFieldNameCharacters As Boolean) As String
        If txtPassword.Text.Contains(";") Then
            Throw New System.Exception("Although Quick Base allows semicolons in passwords the ODBC standard does not permit semicolons." & vbCrLf & "Please change your Quick Base password to eliminate semicolons or use a Quick Base user token instead of a password.")
            Return ""
        End If

        Dim connectionString As String = "Driver={QuNect ODBC for QuickBase};CSVCHUNKSIZE=1000000;uid=" & txtUsername.Text & ";pwd=" & txtPassword.Text & ";QUICKBASESERVER=" & txtServer.Text & ";APPTOKEN=" & txtAppToken.Text
        If usefids Then
            connectionString &= ";USEFIDS=1"
        End If
        If allFieldNameCharacters Then
            connectionString &= ";FIELDNAMECHARACTERS=all"
        End If
        If cmbPassword.SelectedIndex = PasswordOrToken.Neither Then
            cmbPassword.Focus()
            Throw New System.Exception("Please indicate whether you are using a password or a user token.")
            Return ""
        ElseIf cmbPassword.SelectedIndex = PasswordOrToken.password Then
            connectionString &= ";PWDISPASSWORD=1"
        Else
            connectionString &= ";PWDISPASSWORD=0"
        End If
#If DEBUG Then
        'connectionString &= ";LOGSQL=1;LOGAPI=1"
#End If
        If lblDebug.Visible Then
            connectionString &= ";LOGSQL=1"
        End If
        Return connectionString
    End Function
    Private Function getHashSetofFieldValues(dbid As String, fid As String) As HashSet(Of String)
        Dim strSQL As String = "SELECT fid" & fid & " FROM " & dbid
        getHashSetofFieldValues = New HashSet(Of String)
        Dim connectionString As String = getConnectionString(True, False) '"Driver={QuNect ODBC For QuickBase};uid=" & txtUsername.Text & ";pwd=" & txtPassword.Text & ";QUICKBASESERVER=" & txtServer.Text & ";USEFIDS=1;APPTOKEN=" & txtAppToken.Text
        Using quNectConn As OdbcConnection = getquNectConn(connectionString)
            Using quNectCmd As OdbcCommand = New OdbcCommand(strSQL, quNectConn)
                Dim dr As OdbcDataReader
                Try
                    dr = quNectCmd.ExecuteReader()
                    If Not dr.HasRows Then
                        Exit Function
                    End If
                Catch excpt As Exception
                    Exit Function
                End Try
                While (dr.Read())
                    getHashSetofFieldValues.Add(dr.GetValue(0))
                End While
            End Using
        End Using
    End Function
    Function buildDestinationFieldListAndAddParameters(fids As ArrayList, fieldNodes As Dictionary(Of String, fieldStruct), command As OdbcCommand, BuiltInRecordIDisKey As Boolean) As String

        buildDestinationFieldListAndAddParameters = ""
        Dim qdbType As OdbcType
        For Each fid As String In fids
            qdbType = getODBCTypeFromQuickBaseFieldNode(fieldNodes(fid))
            buildDestinationFieldListAndAddParameters &= "fid" & fid & ","
            If Not command Is Nothing Then
                command.Parameters.Add("@fid" & fid, qdbType)
            End If
        Next
        buildDestinationFieldListAndAddParameters = buildDestinationFieldListAndAddParameters.TrimEnd(",")
    End Function
    Private Function restoreTable(checkForErrorsOnly As Boolean, previewOnly As Boolean, bulkRestore As Boolean) As Boolean
        restoreTable = True
        Dim processOrderFoMappingRows = New ArrayList
        Dim dsPreview As New DataSet
        Dim dtPreview As New DataTable
        Dim drPreview As DataRow = Nothing
        Dim m As Match = Regex.Match(lblFile.Text, "\.csv$", RegexOptions.IgnoreCase)
        If Not m.Success Then
            PopUpMsgBox("Cannot restore from a non CSV file: " & lblFile.Text, MsgBoxStyle.OkOnly, AppName)
            restoreTable = False
            Exit Function
        End If
        Dim dbid As String = Regex.Replace(lblTable.Text, "^.* ", "")

        Dim BuiltInRecordIDisKey As Boolean = False
        If Not checkForErrorsOnly Then
            BuiltInRecordIDisKey = preparePotentialParent(dbid)
        End If

        'need to create the SQL statement
        Dim fids As New ArrayList
        Dim vals As New ArrayList
        Dim importThese As New ArrayList
        Dim fields As New ArrayList
        Dim fieldLabels As New ArrayList
        Dim fidsForImport As New HashSet(Of String)
        Dim requiredFidsForImport As New HashSet(Of String)
        Dim uniqueFidsForImport As New HashSet(Of String)
        Dim conversionErrors As String = ""
        Dim missingRIDs As String = ""
        Dim improperlyFormattedLines As String = ""
        Dim requiredFieldsErrors As String = ""
        Dim uniqueFieldErrors As String = ""
        Dim ridIsMapped As Boolean = False
        Dim keyIsMapped As Boolean = False
        Dim columnsMapped As Integer = 0
        For i As Integer = 0 To dgMapping.Rows.Count - 1
            Dim destComboBoxCell As DataGridViewComboBoxCell = DirectCast(dgMapping.Rows(i).Cells(mapping.destination), System.Windows.Forms.DataGridViewComboBoxCell)
            If destComboBoxCell.Value Is Nothing Then Continue For
            Dim destDDIndex = destComboBoxCell.Items.IndexOf(destComboBoxCell.Value)

            If destDDIndex > 1 Then
                'this is a field that needs importing
                Dim fieldNode As fieldStruct = fieldNodes(destinationLabelsToFids(destComboBoxCell.Value))
                Dim label As String = fieldNode.label
                If fieldNode.parentFieldID <> "" Then
                    label = fieldNodes(fieldNode.parentFieldID).label & ": " & label
                End If
                If previewOnly Then
                    dtPreview.Columns.Add(New DataColumn(label, Type.GetType("System.String")))
                End If
                fields.Add(fieldNode)
                Dim fid As String = fieldNode.fid
                If fidsForImport.Contains(fid) Then
                    PopUpMsgBox("You cannot import two different columns into the same field: " & destComboBoxCell.Value, MsgBoxStyle.OkOnly, AppName)
                    restoreTable = False
                    Exit Function
                End If
                fidsForImport.Add(fid)
                If fid = "3" Then ridIsMapped = True
                If fid = keyfid Then
                    processOrderFoMappingRows.Insert(0, columnsMapped)
                    keyIsMapped = True
                Else
                    processOrderFoMappingRows.Add(columnsMapped)
                End If
                fids.Add(fid)
                vals.Add("?")
                importThese.Add(i)
                fieldLabels.Add(label)
                columnsMapped += 1
            End If
        Next
        'now we need to map the Record ID# field to the "QuNect Restore Temporary Record ID#" field
        If BuiltInRecordIDisKey Then
            Dim fieldNode As New fieldStruct
            fieldNode.label = fieldNameToContainOldRecordIdNumbers
            fieldNode.fid = getFIDofQuNectRestoreTemporaryRecordID(dbid)
            fieldNode.type = "float"
            fieldNode.base_type = "float"
            fieldNodes.Add(fieldNode.fid, fieldNode)
            processOrderFoMappingRows.Add(columnsMapped)
            fids.Add(fieldNode.fid)
            vals.Add("?")
            importThese.Add(importThese(0))
            fieldLabels.Add(fieldNode.label)
        End If
        Dim unmappedRequireds As Boolean = False
        Dim unmappedUniques As Boolean = False
        Dim connectionString As String = getConnectionString(False, False) '"Driver={QuNect ODBC For QuickBase};uid=" & txtUsername.Text & ";pwd=" & txtPassword.Text & ";QUICKBASESERVER=" & txtServer.Text & ";USEFIDS=1;APPTOKEN=" & txtAppToken.Text
        Try
            Using connection As OdbcConnection = getquNectConn(connectionString)
                If connection Is Nothing Then Exit Function
                'also need to pull all the fields marked required
                If checkForErrorsOnly Then
                    frmErr.rdbCancel.Checked = True
                    uniqueExistingFieldValues = New Dictionary(Of String, HashSet(Of String))
                    uniqueNewFieldValues = New Dictionary(Of String, HashSet(Of String))

                    For Each field As KeyValuePair(Of String, fieldStruct) In fieldNodes
                        Dim fid As String = field.Value.fid
                        If field.Value.unique AndAlso field.Value.fid <> "3" Then
                            If Not fidsForImport.Contains(fid) Then
                                If Not keyIsMapped Then
                                    uniqueFieldErrors &= "You must import values into the unique field: " & field.Value.label
                                    restoreTable = False
                                Else
                                    unmappedUniques = True
                                End If
                            ElseIf fid <> keyfid Then
                                'we also need to keep a list of the unique fields that we're importing into to check and make sure all imported values are unique
                                uniqueFidsForImport.Add(fid)
                                'do we need this hashset to check the imported values? We would need to check to see if we were updating an existing record or creating a new record
                                'if we have a 
                                If Not ridIsMapped Then
                                    uniqueExistingFieldValues.Add(fid, getHashSetofFieldValues(dbid, fid))
                                End If
                            End If
                        End If
                        If field.Value.required Then
                            If Not fidsForImport.Contains(fid) Then
                                If Not keyIsMapped Then
                                    requiredFieldsErrors &= "You must import values into the required field: " & field.Value.label
                                    restoreTable = False
                                Else
                                    unmappedRequireds = True
                                End If
                            Else
                                'we also need to keep a list of the required fields that we're importing into
                                requiredFidsForImport.Add(fid)
                            End If
                        End If
                    Next
                End If
                If ridIsMapped And (checkForErrorsOnly Or bulkRestore) Then
                    'need to pull in all the rids from QuickBase
                    rids = getHashSetofFieldValues(dbid, "3")
                End If
                Dim strSQL As String = "INSERT INTO " & dbid & " ("

                strSQL &= buildDestinationFieldListAndAddParameters(fids, fieldNodes, Nothing, BuiltInRecordIDisKey)
                strSQL &= ") VALUES ("
                strSQL &= String.Join(",", vals.ToArray)
                strSQL &= ")"
                Using command As OdbcCommand = New OdbcCommand(strSQL, connection)
                    buildDestinationFieldListAndAddParameters(fids, fieldNodes, command, BuiltInRecordIDisKey)
                    Dim csvReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(lblFile.Text)
                    csvReader.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited
                    csvReader.Delimiters = New String() {","}
                    csvReader.HasFieldsEnclosedInQuotes = True
                    Dim currentRow As String()
                    'Loop through all of the fields in the file. 
                    Dim transaction As OdbcTransaction = Nothing
                    transaction = connection.BeginTransaction()
                    command.Transaction = transaction
                    command.CommandType = CommandType.Text
                    command.CommandTimeout = 0
                    Dim lineCounter As Integer = 0
                    Dim fileLineCounter As Integer = 0
                    If Not csvReader.EndOfData And chkBxHeaders.Checked Then
                        currentRow = csvReader.ReadFields()
                        fileLineCounter = 1
                    End If
                    Try
                        pb.Value = 0
                        pb.Visible = True
                        pb.Maximum = 1000
                        While Not csvReader.EndOfData
                            Try
                                lblProgress.Text = "Reading CSV line " & fileLineCounter & " found " & lineCounter & " records for import."
                                If previewOnly Then
                                    lblMode.Text = "Applying import criteria and checking for errors to display a preview of the import."
                                ElseIf checkForErrorsOnly Then
                                    lblMode.Text = "Applying import criteria and checking for errors in advance of importing."
                                Else
                                    lblMode.Text = "Applying import criteria and importing."
                                End If

                                fileLineCounter += 1
                                Application.DoEvents()
                                currentRow = csvReader.ReadFields()
                                'here we need to walk down the rows of the criteria datagrid to see if we need to exclude this row
                                Dim skipRow As Boolean = False
                                For i As Integer = 0 To dgCriteria.Rows.Count - 1
                                    Dim sourceLabel As String = DirectCast(dgCriteria.Rows(i).Cells(mapping.source), System.Windows.Forms.DataGridViewComboBoxCell).Value
                                    If sourceLabel Is Nothing Then Continue For
                                    Dim fieldType As String = sourceLabelToFieldType(sourceLabel)
                                    Dim cellValue As String = currentRow(sourceFieldNames(sourceLabel))
                                    Dim op As String = DirectCast(dgCriteria.Rows(i).Cells(filter.booleanOperator), System.Windows.Forms.DataGridViewComboBoxCell).Value
                                    Dim criteria As String = DirectCast(dgCriteria.Rows(i).Cells(filter.criteria), System.Windows.Forms.DataGridViewTextBoxCell).Value
                                    If skipThisRow(op, cellValue, criteria, fieldType) Then
                                        skipRow = True
                                        Exit For
                                    End If
                                Next
                                If skipRow Then
                                    Continue While
                                End If
                                'This is a row we need to import
                                If previewOnly Then
                                    drPreview = dtPreview.NewRow()
                                End If
                                Dim importingIntoExistingRecord As Boolean = False
                                For k As Integer = 0 To processOrderFoMappingRows.Count - 1
                                    Dim i As Integer = processOrderFoMappingRows(k)
                                    Dim val As String = currentRow(importThese(i))
                                    If previewOnly Then
                                        drPreview(i) = val
                                    End If
                                    If checkForErrorsOnly Then
                                        If keyIsMapped AndAlso k = 0 Then
                                            'check if this is a new or existing key value
                                            If fids(i) = "3" Then
                                                If Not rids.Contains(val) Then
                                                    'here we're going to have a problem because this record no longer exists in QuickBase
                                                    'we could update all the child tables with the newly minted Record ID#s
                                                    'but we wouid have to do one record at a time to accomplish this

                                                    missingRIDs &= vbCrLf & "line " & fileLineCounter & " has Record ID# " & val & " which no longer exists"
                                                    If missingRIDs.Length > 1000 Then
                                                        missingRIDs &= vbCrLf & "There may be additional errors beyond the ones above."
                                                        Exit While
                                                    End If
                                                Else
                                                    importingIntoExistingRecord = True
                                                End If
                                            Else
                                                'the key field is not Record ID#
                                                'does this value exist in the table already?
                                                If uniqueExistingFieldValues(fids(i)).Contains(val) Then
                                                    importingIntoExistingRecord = True
                                                End If
                                            End If
                                        End If
                                        If requiredFidsForImport.Contains(fids(i)) Then
                                            If val = "" Then
                                                requiredFieldsErrors &= vbCrLf & "line " & fileLineCounter & " has a blank value mapped to required field " & fieldLabels(i)
                                                If requiredFieldsErrors.Length > 1000 Then
                                                    requiredFieldsErrors &= vbCrLf & "There may be additional errors beyond the ones above."
                                                    Exit While
                                                End If
                                            End If
                                        End If
                                        If uniqueFidsForImport.Contains(fids(i)) Then
                                            'here we need to see if we have a duplicate
                                            If Not uniqueNewFieldValues.Contains(fids(i)) Then
                                                uniqueNewFieldValues.Add(fids(i), New HashSet(Of String))
                                            End If
                                            If uniqueNewFieldValues(fids(i)).Contains(val) Then
                                                uniqueFieldErrors &= vbCrLf & "line " & fileLineCounter & " has a duplicate value for unique field " & fieldLabels(i)
                                            Else
                                                uniqueNewFieldValues(fids(i)).Add(val)
                                            End If
                                            If val = "" Then
                                                uniqueFieldErrors &= vbCrLf & "line " & fileLineCounter & " has a blank value mapped to unique field " & fieldLabels(i)
                                            End If
                                            If uniqueFieldErrors.Length > 1000 Then
                                                uniqueFieldErrors &= vbCrLf & "There may be additional errors beyond the ones above."
                                                Exit While
                                            End If
                                        End If

                                    ElseIf fids(i) = "3" AndAlso Not rids.Contains(val) Then
                                        If frmErr.rdbSkipRecords.Checked Then
                                            Continue While
                                        Else
                                            val = ""
                                        End If
                                    End If

                                    If setODBCParameter(val, fids(i), fieldLabels(i), command, fileLineCounter, checkForErrorsOnly, conversionErrors) Then
                                        Exit While
                                    End If

                                Next
                                If Not importingIntoExistingRecord Then
                                    If unmappedRequireds Then
                                        requiredFieldsErrors &= vbCrLf & "line " & fileLineCounter & " will create a new record and leave required fields blank"
                                    End If
                                    If unmappedUniques Then
                                        uniqueFieldErrors &= vbCrLf & "line " & fileLineCounter & " will create a new record and leave unique fields blank"
                                    End If
                                End If
                                If previewOnly Then
                                    dtPreview.Rows.Add(drPreview)
                                End If
                                If Not checkForErrorsOnly Then command.ExecuteNonQuery()

                            Catch ex As Exception
                                If checkForErrorsOnly Then
                                    improperlyFormattedLines &= vbCrLf & "line " & fileLineCounter & " Is Not a properly formatted CSV line."
                                    If improperlyFormattedLines.Length > 1000 Then
                                        improperlyFormattedLines &= vbCrLf & "There may be additional errors beyond the ones above."
                                        Exit While
                                    End If
                                Else
                                    Throw New Exception("line " & fileLineCounter & " in '" & lblFile.Text & "' has an issue. " & ex.Message)
                                End If
                            End Try
                            lineCounter += 1
                            pb.Value = fileLineCounter Mod 1000
                        End While
                        If Not checkForErrorsOnly Then
                            transaction.Commit()
                            If Not bulkRestore Then
                                PopUpMsgBox("Imported " & lineCounter & " records!", MsgBoxStyle.OkOnly, AppName)
                            End If
                        Else
                            restoreTable = showErrors(previewOnly, conversionErrors, improperlyFormattedLines, uniqueFieldErrors, requiredFieldsErrors, missingRIDs)
                            pb.Visible = False
                        End If
                        If previewOnly And restoreTable Then
                            dsPreview.Tables.Add(dtPreview)
                            frmPreview.dgPreview.DataSource = dsPreview.Tables(0)
                            frmPreview.lblPreview.Text = "Out of " & fileLineCounter & " lines in CSV found " & lineCounter & " records for import."
                            frmPreview.ShowDialog()
                        End If
                    Catch ex As Exception
                        PopUpMsgBox("Could Not import because " & ex.Message & " " & strSQL)
                        restoreTable = False
                    Finally
                        lblProgress.Text = ""
                        pb.Visible = False
                    End Try
                End Using
            End Using
        Catch ex As Exception
            restoreTable = False
        End Try
    End Function
    Private Function getquNectConn(connectionString As String) As OdbcConnection

        Dim quNectConn As OdbcConnection = New OdbcConnection(connectionString)
        Try
            quNectConn.Open()
        Catch excpt As Exception
            Me.Cursor = Cursors.Default
            If excpt.Message.StartsWith("Error [IM003]") Or excpt.Message.Contains("Data source name Not found") Then
                Throw New Exception("Please install QuNect ODBC For QuickBase from http://qunect.com/download/QuNect.exe and try again.")
            Else
                Throw New Exception(excpt.Message.Substring(13))
            End If
            Return Nothing
            Exit Function
        End Try

        Dim ver As String = quNectConn.ServerVersion
        Dim m As Match = Regex.Match(ver, "\d+\.(\d+)\.(\d+)\.(\d+)")
        qdbVer.year = CInt(m.Groups(1).Value)
        qdbVer.major = CInt(m.Groups(2).Value)
        qdbVer.minor = CInt(m.Groups(3).Value)
        If (qdbVer.major < 7) Or (qdbVer.major = 7 And qdbVer.minor < 84) Then
            Throw New Exception("You are running the " & ver & " version of QuNect ODBC for QuickBase. Please install the latest version from http://qunect.com/download/QuNect.exe")
            quNectConn.Close()
            Me.Cursor = Cursors.Default
            Return Nothing
            Exit Function
        End If
        Return quNectConn
    End Function
    Private Function showErrors(previewOnly As Boolean, ByRef conversionErrors As String, ByRef improperlyFormattedLines As String, ByRef uniqueFieldErrors As String, ByRef requiredFieldsErrors As String, ByRef missingRIDs As String) As Boolean
        If conversionErrors <> "" Or improperlyFormattedLines <> "" Or uniqueFieldErrors <> "" Or requiredFieldsErrors <> "" Or missingRIDs <> "" Then
            frmErr.lblMalformed.Text = improperlyFormattedLines
            frmErr.lblUnique.Text = uniqueFieldErrors
            frmErr.lblRequired.Text = requiredFieldsErrors
            frmErr.lblConversions.Text = conversionErrors
            frmErr.lblMissing.Text = missingRIDs
            If improperlyFormattedLines <> "" Then
                frmErr.TabControlErrors.SelectedIndex = errorTabs.malformed
                frmErr.TabMalformed.Text = "*Malformed lines"
            Else
                frmErr.TabMalformed.Text = "Malformed lines"
            End If
            If uniqueFieldErrors <> "" Then
                frmErr.TabControlErrors.SelectedIndex = errorTabs.unique
                frmErr.TabUnique.Text = "*Unique"
            Else
                frmErr.TabUnique.Text = "Unique"
            End If
            If requiredFieldsErrors <> "" Then
                frmErr.TabControlErrors.SelectedIndex = errorTabs.required
                frmErr.TabRequired.Text = "*Required"
            Else
                frmErr.TabRequired.Text = "Required"
            End If
            If conversionErrors <> "" Then
                frmErr.TabControlErrors.SelectedIndex = errorTabs.conversions
                frmErr.TabConversions.Text = "*Conversion"
            Else
                frmErr.TabConversions.Text = "Conversion"
            End If
            If missingRIDs <> "" Then
                frmErr.TabControlErrors.SelectedIndex = errorTabs.missing
                frmErr.TabMissing.Text = "*Missing Record ID#s"
            Else
                frmErr.TabMissing.Text = "Missing Record ID#s"
            End If
            If previewOnly Then
                frmErr.pnlButtons.Visible = False
            Else
                frmErr.pnlButtons.Visible = True
            End If
            Dim diagResult As DialogResult = frmErr.ShowDialog()
            If frmErr.rdbCancel.Checked Then
                Return False
            Else
                Return True
            End If
        End If
        Return True
    End Function
    Function setODBCParameter(val As String, fid As String, fieldLabel As String, command As OdbcCommand, ByRef fileLineCounter As Integer, checkForErrorsOnly As Boolean, ByRef conversionErrors As String) As Boolean
        Dim qdbType As OdbcType = command.Parameters("@fid" & fid).OdbcType
        Try
            Select Case qdbType
                Case OdbcType.Int
                    If val = "" Then
                        Dim nullInteger As Integer
                        command.Parameters("@fid" & fid).Value = nullInteger
                    Else
                        command.Parameters("@fid" & fid).Value = Convert.ToInt32(val)
                    End If
                Case OdbcType.Double, OdbcType.Numeric
                    command.Parameters("@fid" & fid).Value = Convert.ToDouble(val)
                Case OdbcType.Date
                    command.Parameters("@fid" & fid).Value = Date.Parse(val)
                Case OdbcType.DateTime
                    command.Parameters("@fid" & fid).Value = DateTime.Parse(val)
                Case OdbcType.Time
                    command.Parameters("@fid" & fid).Value = TimeSpan.Parse(val)
                Case OdbcType.Bit
                    Dim match As Match = isBooleanTrue.Match(val)
                    If match.Success Then
                        command.Parameters("@fid" & fid).Value = True
                    Else
                        command.Parameters("@fid" & fid).Value = False
                    End If
                Case Else
                    command.Parameters("@fid" & fid).Value = val
            End Select
        Catch ex As Exception
            If checkForErrorsOnly And val <> "" Then
                conversionErrors &= vbCrLf & "line " & fileLineCounter + 1 & " '" & val & "' to " & qdbType.ToString() & " for field " & fieldLabel
                If conversionErrors.Length > 1000 Then
                    conversionErrors &= vbCrLf & "There may be additional errors beyond the ones above."
                    Return True
                End If
            End If
            If frmErr.rdbSkipRecords.Checked Then
                fileLineCounter += 1
                Return False
            End If
            If Not checkForErrorsOnly Then
                Select Case qdbType
                    Case OdbcType.Int, OdbcType.Double, OdbcType.Numeric
                        Dim nullDouble As Double
                        command.Parameters("@fid" & fid).Value = nullDouble
                    Case OdbcType.Date
                        Dim nulldate As Date
                        command.Parameters("@fid" & fid).Value = nulldate
                    Case OdbcType.DateTime
                        Dim nullDateTime As DateTime
                        command.Parameters("@fid" & fid).Value = nullDateTime
                    Case OdbcType.Time
                        Dim nullTimeSpan As TimeSpan
                        command.Parameters("@fid" & fid).Value = nullTimeSpan
                    Case OdbcType.Bit
                        command.Parameters("@fid" & fid).Value = False
                    Case Else
                        command.Parameters("@fid" & fid).Value = ""
                End Select
            End If
        End Try
        Return False
    End Function
    Private Function getODBCTypeFromQuickBaseFieldNode(fieldNode As fieldStruct) As OdbcType
        Select Case fieldNode.base_type
            Case "text"
                Return OdbcType.VarChar
            Case "float"
                If fieldNode.decimal_places > 0 Then
                    Return OdbcType.Numeric
                Else
                    Return OdbcType.Double
                End If

            Case "bool"
                Return OdbcType.Bit
            Case "int32"
                Return OdbcType.Int
            Case "int64"
                Select Case fieldNode.type
                    Case "timestamp"
                        Return OdbcType.DateTime
                    Case "timeofday"
                        Return OdbcType.Time
                    Case "duration"
                        Return OdbcType.Double
                    Case Else
                        Return OdbcType.Date
                End Select
            Case Else
                Return OdbcType.VarChar
        End Select
        Return OdbcType.VarChar
    End Function

    Private Function skipThisRow(op As String, cellValue As String, criteria As String, fieldType As String) As Boolean
        If op = "is null" AndAlso cellValue = "" Then
            Return False
        ElseIf op = "is null" AndAlso cellValue <> "" Then
            Return True
        ElseIf op = "is not null" AndAlso cellValue <> "" Then
            Return False
        ElseIf op = "is not null" AndAlso cellValue = "" Then
            Return True
        End If
        Dim compareResult As comparisonResult = compare(cellValue, criteria, fieldType)
        If compareResult = comparisonResult.null Then
            Return True
        End If
        Select Case op
            Case "equals"
                If compareResult <> comparisonResult.equal Then
                    Return True
                End If
            Case "does not equal"
                If compareResult = comparisonResult.equal Then
                    Return True
                End If
            Case "greater than or equal"
                If compareResult = comparisonResult.less Then
                    Return True
                End If
            Case "greater than"
                If compareResult <> comparisonResult.greater Then
                    Return True
                End If
            Case "less than or equal"
                If compareResult = comparisonResult.greater Then
                    Return True
                End If
            Case "less than"
                If compareResult <> comparisonResult.less Then
                    Return True
                End If
            Case "contains"
                If Not cellValue.Contains(criteria) Then
                    Return True
                End If
            Case "starts with"
                If Not cellValue.StartsWith(criteria) Then
                    Return True
                End If
            Case "does not contain"
                If cellValue.Contains(criteria) Then
                    Return True
                End If
            Case "does not start with"
                If cellValue.StartsWith(criteria) Then
                    Return True
                End If
            Case Else
                Return False
        End Select
        Return False
    End Function

    Private Function compare(leftStr As String, rightStr As String, fieldType As String) As comparisonResult
        Dim rightVar As Object = 0
        Dim leftVar As Object = 0
        Try
            Select Case fieldType
                Case "Date"
                    rightVar = Date.Parse(rightStr)
                    leftVar = Date.Parse(leftStr)
                    If leftVar = rightVar Then
                        Return comparisonResult.equal
                    ElseIf Date.Compare(leftVar, rightVar) > 0 Then
                        Return comparisonResult.greater
                    Else
                        Return comparisonResult.less
                    End If
                Case "timestamp"
                    rightVar = DateTime.Parse(rightStr)
                    leftVar = DateTime.Parse(leftStr)
                    If leftVar = rightVar Then
                        Return comparisonResult.equal
                    ElseIf DateTime.Compare(leftVar, rightVar) > 0 Then
                        Return comparisonResult.greater
                    Else
                        Return comparisonResult.less
                    End If
                Case "float", "currency"
                    rightVar = Convert.ToDouble(rightStr)
                    leftVar = Convert.ToDouble(leftStr)
                    If leftVar = rightVar Then
                        Return comparisonResult.equal
                    ElseIf leftVar > rightVar Then
                        Return comparisonResult.greater
                    Else
                        Return comparisonResult.less
                    End If
                Case "recordid"
                    rightVar = Convert.ToInt32(rightStr)
                    leftVar = Convert.ToInt32(leftStr)
                    If leftVar = rightVar Then
                        Return comparisonResult.equal
                    ElseIf leftVar > rightVar Then
                        Return comparisonResult.greater
                    Else
                        Return comparisonResult.less
                    End If
                Case "checkbox"
                    Dim match As Match = isBooleanTrue.Match(leftStr)
                    If match.Success Then
                        leftVar = True
                    Else
                        leftVar = False
                    End If
                    match = isBooleanTrue.Match(rightStr)
                    If match.Success Then
                        rightVar = True
                    Else
                        rightVar = False
                    End If
                    If leftVar = rightVar Then
                        Return comparisonResult.equal
                    Else
                        Return comparisonResult.notEqual
                    End If
                    'need to deal with duration and timeofday
                Case "duration", "timeofday"
                    Dim tsLeft As TimeSpan
                    Dim tsRight As TimeSpan
                    If TimeSpan.TryParse(rightStr, tsRight) Then
                        tsLeft = TimeSpan.FromMilliseconds(Convert.ToDouble(leftStr))
                    Else
                        Return comparisonResult.null
                    End If
                    If tsLeft = tsRight Then
                        Return comparisonResult.equal
                    ElseIf tsLeft > tsRight Then
                        Return comparisonResult.greater
                    Else
                        Return comparisonResult.less
                    End If
                Case Else
                    If leftStr <> rightStr Then
                        Return comparisonResult.notEqual
                    Else
                        Return comparisonResult.equal
                    End If
            End Select
        Catch ex As Exception
            Return comparisonResult.null
        End Try
    End Function
    Sub showHideControls()
        cmbPassword.Visible = txtUsername.Text.Length > 0
        txtPassword.Visible = cmbPassword.SelectedIndex > PasswordOrToken.Neither
        txtServer.Visible = txtUsername.Text.Length > 0 And txtPassword.Text.Length > 0 And cmbPassword.SelectedIndex > PasswordOrToken.Neither
        lblServer.Visible = txtServer.Visible
        cmbBulkorSingle.Visible = txtUsername.Text.Length > 0 And cmbPassword.SelectedIndex > PasswordOrToken.Neither And txtPassword.Text.Length > 0 And txtServer.Text.Length > 0
        txtAppToken.Visible = txtUsername.Text.Length > 0 And cmbPassword.SelectedIndex = PasswordOrToken.password And txtPassword.Text.Length > 0 And txtServer.Text.Length > 0
        lblAppToken.Visible = txtAppToken.Visible
        btnAppToken.Visible = txtAppToken.Visible
        btnUserToken.Visible = txtUsername.Text.Length > 0 And cmbPassword.SelectedIndex = PasswordOrToken.token
        btnListTables.Visible = txtUsername.Text.Length > 0 And cmbPassword.SelectedIndex > PasswordOrToken.Neither And txtPassword.Text.Length > 0 And txtServer.Text.Length > 0 And cmbBulkorSingle.SelectedIndex = BulkOrSingle.One And lblFile.Text.Length > 0
        lblTable.Visible = btnListTables.Visible
        btnSource.Visible = cmbBulkorSingle.Visible And cmbBulkorSingle.SelectedIndex > BulkOrSingle.Neither
        lblFile.Visible = cmbBulkorSingle.Visible And cmbBulkorSingle.SelectedIndex > BulkOrSingle.Neither
        ckbDetectProxy.Visible = txtServer.Text.Length > 0 And txtServer.Visible
        chkBxHeaders.Visible = lblFile.Visible And lblFile.Text.Length > 0
        btnPreview.Visible = cmbBulkorSingle.Visible And cmbBulkorSingle.SelectedIndex = BulkOrSingle.One And lblTable.Text.Length > 0
        btnImport.Visible = cmbBulkorSingle.Visible And ((cmbBulkorSingle.SelectedIndex > BulkOrSingle.One And lblTable.Text.Length > 0) Or btnPreview.Visible)
        dgCriteria.Visible = btnPreview.Visible
        dgMapping.Visible = btnPreview.Visible
        If cmbBulkorSingle.SelectedIndex = BulkOrSingle.One Then
            btnSource.Text = "Choose file to import"
            btnImport.Text = "Import from CSV file"
        ElseIf cmbBulkorSingle.SelectedIndex = BulkOrSingle.TwoOrMore Then
            btnSource.Text = "Choose folder to import"
            btnImport.Text = "Import from folder of CSV files"
        End If
    End Sub

    Private Sub txtUsername_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUsername.TextChanged
        SaveSetting(AppName, "Credentials", "username", txtUsername.Text)
        showHideControls()
    End Sub
    Private Sub cmbPassword_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbPassword.SelectedIndexChanged
        SaveSetting(AppName, "Credentials", "passwordOrToken", cmbPassword.SelectedIndex)
        If lblTable.Text.Length > 0 And cmbBulkorSingle.SelectedIndex = BulkOrSingle.One Then
            listFieldsAndReturnKeyFID(lblTable.Text, False)
        End If
        showHideControls()
    End Sub
    Private Sub txtPassword_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPassword.TextChanged
        SaveSetting(AppName, "Credentials", "password", txtPassword.Text)
        showHideControls()
    End Sub
    Private Sub txtServer_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtServer.TextChanged
        SaveSetting(AppName, "Credentials", "server", txtServer.Text)
        showHideControls()
    End Sub
    Private Sub cmbBulkorSingle_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbBulkorSingle.SelectedIndexChanged
        SaveSetting(AppName, "config", "bulkOrSingle", cmbBulkorSingle.SelectedIndex)
        showHideControls()
    End Sub
    Private Sub txtAppToken_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAppToken.TextChanged
        SaveSetting(AppName, "Credentials", "apptoken", txtAppToken.Text)
    End Sub
    Private Sub lblTable_TextChanged(sender As Object, e As EventArgs) Handles lblTable.TextChanged
        SaveSetting(AppName, "config", "table", lblTable.Text)
        showHideControls()
        If lblTable.Text.Length > 0 And cmbBulkorSingle.SelectedIndex = BulkOrSingle.One Then
            listFieldsAndReturnKeyFID(lblTable.Text, False)
        End If
    End Sub

    Private Sub lblFile_TextChanged(sender As Object, e As EventArgs) Handles lblFile.TextChanged
        SaveSetting(AppName, "config", "file", lblFile.Text)
        showHideControls()
    End Sub

    Private Sub btnSource_Click(sender As Object, e As EventArgs) Handles btnSource.Click
        If cmbBulkorSingle.SelectedIndex = BulkOrSingle.TwoOrMore Then
            If (FolderBrowserDialog.ShowDialog() = DialogResult.OK) Then
                lblFile.Text = FolderBrowserDialog.SelectedPath.ToString()
                btnImport.Visible = True
            End If
        Else
            OpenSourceFile.ShowDialog()
        End If
    End Sub

    Private Sub OpenSourceFile_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenSourceFile.FileOk
        lblFile.Text = OpenSourceFile.FileName.ToString()
        listFieldsAndReturnKeyFID(lblTable.Text, False)
    End Sub

    Private Sub btnListTables_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnListTables.Click
        listTables()
    End Sub
    Private Sub listTables()
        Me.Cursor = Cursors.WaitCursor
        Try
            Dim connectionString As String = getConnectionString(False, True) '"Driver={QuNect ODBC for QuickBase};FIELDNAMECHARACTERS=all;uid=" & txtUsername.Text & ";pwd=" & txtPassword.Text & ";QUICKBASESERVER=" & txtServer.Text & ";APPTOKEN=" & txtAppToken.Text
            Using quNectConn As OdbcConnection = getquNectConn(connectionString)
                Dim tables As DataTable = quNectConn.GetSchema("Tables")
                listTablesFromGetSchema(tables)
                quNectConn.Close()
                quNectConn.Dispose()
            End Using
        Catch ex As Exception
            Me.Cursor = Cursors.Default
            PopUpMsgBox(ex.Message)
        End Try
    End Sub
    Sub listTablesFromGetSchema(tables As DataTable)
        frmTableChooser.tvAppsTables.BeginUpdate()
        frmTableChooser.tvAppsTables.Nodes.Clear()
        frmTableChooser.tvAppsTables.ShowNodeToolTips = True
        Dim dbName As String
        Dim applicationName As String = ""
        Dim prevAppName As String = ""
        Dim dbid As String
        pb.Value = 0
        pb.Visible = True
        pb.Maximum = tables.Rows.Count


        For i = 0 To tables.Rows.Count - 1
            pb.Value = i
            Application.DoEvents()
            dbName = tables.Rows(i)(2)
            applicationName = tables.Rows(i)(0)
            dbid = tableNameToDBID(dbName)
            If applicationName <> prevAppName Then

                Dim appNode As TreeNode = frmTableChooser.tvAppsTables.Nodes.Add(applicationName)
                appNode.Tag = dbid
                prevAppName = applicationName
            End If
            Dim tableName As String = dbName
            If dbName.Length > applicationName.Length Then
                tableName = dbName.Substring(applicationName.Length + 2)
            End If
            Dim tableNode As TreeNode = frmTableChooser.tvAppsTables.Nodes(frmTableChooser.tvAppsTables.Nodes.Count - 1).Nodes.Add(tableName)

            tableNode.Tag = dbid

        Next

        pb.Visible = False
        frmTableChooser.tvAppsTables.EndUpdate()
        pb.Value = 0
        btnImport.Visible = True
        lblTable.Visible = True
        frmTableChooser.Show()
        Me.Cursor = Cursors.Default
    End Sub
    Function getFIDofQuNectRestoreTemporaryRecordID(dbid As String) As String
        getFIDofQuNectRestoreTemporaryRecordID = ""
        Try
            Dim connectionString As String = getConnectionString(False, False) '"Driver={QuNect ODBC for QuickBase};uid=" & txtUsername.Text & ";pwd=" & txtPassword.Text & ";QUICKBASESERVER=" & txtServer.Text & ";APPTOKEN=" & txtAppToken.Text
            Dim strSQL As String = "SELECT fid FROM """ & dbid & "~fields"" WHERE label = 'QuNect Restore Temporary Record ID#'"
            Using connection As OdbcConnection = getquNectConn(connectionString)
                If connection Is Nothing Then Exit Function

                Using quNectCmd As OdbcCommand = New OdbcCommand(strSQL, connection)
                    Dim dr As OdbcDataReader
                    Try
                        dr = quNectCmd.ExecuteReader()
                    Catch excpt As Exception
                        quNectCmd.Dispose()
                        PopUpMsgBox(excpt.Message, MsgBoxStyle.OkOnly, AppName)
                        Exit Function
                    End Try
                    If Not dr.HasRows Then
                        Exit Function
                    End If
                    dr.Read()
                    getFIDofQuNectRestoreTemporaryRecordID = dr.GetString(0)
                    quNectCmd.Dispose()
                End Using
            End Using
        Catch ex As Exception
            PopUpMsgBox("Could not get fid of 'QuNect Restore Temporary Record ID#' in " & dbid & " " & ex.Message)
        End Try
    End Function
    Function preparePotentialParent(dbid As String) As Boolean
        preparePotentialParent = False
        Try
            Dim connectionString As String = getConnectionString(False, False) '"Driver={QuNect ODBC for QuickBase};uid=" & txtUsername.Text & ";pwd=" & txtPassword.Text & ";QUICKBASESERVER=" & txtServer.Text & ";APPTOKEN=" & txtAppToken.Text
            Dim strSQL As String = "SELECT label, fid, iskey FROM """ & dbid & "~fields"" WHERE (mode = '' and role = '') or fid = '3'"
            Dim keyFID As String = ""
            Dim keyField As New fieldStruct
            Dim builtInRecordIDLabel As String = ""
            Using connection As OdbcConnection = getquNectConn(connectionString)
                If connection Is Nothing Then Exit Function

                Using quNectCmd As OdbcCommand = New OdbcCommand(strSQL, connection)
                    Dim dr As OdbcDataReader
                    Try
                        dr = quNectCmd.ExecuteReader()
                    Catch excpt As Exception
                        quNectCmd.Dispose()
                        PopUpMsgBox(excpt.Message, MsgBoxStyle.OkOnly, AppName)
                        Exit Function
                    End Try
                    If Not dr.HasRows Then
                        Exit Function
                    End If
                    While (dr.Read())
                        Dim strFID As String = dr.GetString(1)
                        Dim strLabel As String = dr.GetString(0)
                        If strFID = "3" Then
                            builtInRecordIDLabel = strLabel
                        End If
                        If dr.GetBoolean(2) Then
                            keyField.label = strLabel
                            keyField.fid = strFID
                            keyFID = keyField.fid
                            Exit While
                        End If
                    End While
                    quNectCmd.Dispose()
                End Using
            End Using
            Using connection As OdbcConnection = getquNectConn(connectionString)
                If keyFID = "3" Then
                    'we need to create a shadow key field
                    strSQL = "ALTER TABLE """ & dbid & """ ADD """ & fieldNameToContainOldRecordIdNumbers & """ DOUBLE"
                    Using quNectCmd As OdbcCommand = New OdbcCommand(strSQL, connection)
                        Try
                            quNectCmd.ExecuteNonQuery()
                        Catch ex As Exception
                            If Not ex.Message.Contains("Column label already exists") Then
                                Throw ex
                            End If
                        Finally
                            quNectCmd.Dispose()
                        End Try
                    End Using
                    preparePotentialParent = True
                End If
            End Using
        Catch ex As Exception
            PopUpMsgBox("Could not prepare potential parent " & dbid & " " & ex.Message)
        End Try
    End Function
    Function listFieldsAndReturnKeyFID(dbid As String, bulkRestore As Boolean) As String
        listFieldsAndReturnKeyFID = ""
        sourceLabelToFieldType = New Dictionary(Of String, String)
        destinationLabelsToFids = New Dictionary(Of String, String)
        Dim fidToLabel As New Dictionary(Of String, String)
        fieldNodes = New Dictionary(Of String, fieldStruct)
        sourceFieldNames.Clear()
        Try
            Dim connectionString As String = getConnectionString(False, False) '"Driver={QuNect ODBC For QuickBase};uid=" & txtUsername.Text & ";pwd=" & txtPassword.Text & ";QUICKBASESERVER=" & txtServer.Text & ";APPTOKEN=" & txtAppToken.Text
            Dim currentRow As String()
            Using connection As OdbcConnection = getquNectConn(connectionString)
                If connection Is Nothing Then Exit Function
                Dim strSQL As String = "Select label, fid, field_type, parentFieldID, isunique, required, iskey, base_type, decimal_places  FROM """ & dbid & "~fields"" WHERE (mode = '' and role = '') or fid = '3'"

                Dim quNectCmd As OdbcCommand = New OdbcCommand(strSQL, connection)
                Dim dr As OdbcDataReader
                Try
                    dr = quNectCmd.ExecuteReader()
                Catch excpt As Exception
                    quNectCmd.Dispose()
                    PopUpMsgBox(excpt.Message, MsgBoxStyle.OkOnly, AppName)
                    Exit Function
                End Try
                If Not dr.HasRows Then
                    Exit Function
                End If
                Try
                    Dim fidReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(Regex.Replace(lblFile.Text, "\.csv$", ".fids"))
                    fidReader.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited
                    fidReader.Delimiters = New String() {"."}
                    clist = fidReader.ReadFields()
                    If Not fidReader.EndOfData Then
                        fieldTypes = fidReader.ReadFields()
                    Else
                        fieldTypes = Nothing
                    End If
                Catch
                    fieldTypes = Nothing
                    clist = Nothing
                End Try
                'here we need to open the csv and get the field names
                Dim csvReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(lblFile.Text)
                csvReader.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited
                csvReader.Delimiters = New String() {","}
                csvReader.HasFieldsEnclosedInQuotes = True
                'Loop through all of the fields in the schema.
                DirectCast(dgMapping.Columns(mapping.destination), System.Windows.Forms.DataGridViewComboBoxColumn).Items.Clear()
                DirectCast(dgCriteria.Columns(filter.source), System.Windows.Forms.DataGridViewComboBoxColumn).Items.Clear()
                DirectCast(dgMapping.Columns(mapping.destination), System.Windows.Forms.DataGridViewComboBoxColumn).Items.Add("")
                DirectCast(dgMapping.Columns(mapping.destination), System.Windows.Forms.DataGridViewComboBoxColumn).Items.Add("Do not import")
                While (dr.Read()) 'we are looping through all the fields in the destination table here
                    Dim field As New fieldStruct
                    field.label = dr.GetString(0)
                    If field.label = fieldNameToContainOldRecordIdNumbers Then
                        Continue While
                    End If
                    field.fid = dr.GetString(1)
                    field.type = dr.GetString(2)
                    field.parentFieldID = dr.GetString(3)
                    field.unique = dr.GetBoolean(4)
                    field.required = dr.GetBoolean(5)
                    If fidToLabel.ContainsKey(field.fid) Then
                        Throw New System.Exception("Duplicate fid '" & field.fid & "' in " & dbid)
                    End If
                    fidToLabel.Add(field.fid, field.label)
                    If field.parentFieldID <> "" Then
                        field.label = fidToLabel(field.parentFieldID) & ": " & field.label
                    End If
                    If dr.GetBoolean(6) Then
                        listFieldsAndReturnKeyFID = field.fid
                    End If
                    field.base_type = dr.GetString(7)
                    field.decimal_places = 0
                    If Not IsDBNull(dr(8)) Then
                        field.decimal_places = dr.GetInt32(8)
                    End If
                    fieldNodes.Add(field.fid, field)
                    If destinationLabelsToFids.ContainsKey(field.label) Then
                        Throw New System.Exception("Duplicate field label '" & field.label & "' in Quick Base destination table: " & dbid)
                    End If
                    destinationLabelsToFids.Add(field.label, field.fid)
                    DirectCast(dgMapping.Columns(mapping.destination), System.Windows.Forms.DataGridViewComboBoxColumn).Items.Add(field.label)
                End While
                quNectCmd.Dispose()
                If Not csvReader.EndOfData Then
                    Try
                        currentRow = csvReader.ReadFields()
                        dgMapping.Rows.Clear()
                        Dim j As Integer = 0
                        For i As Integer = 0 To currentRow.Length - 1
                            Dim sourceFieldName As String = currentRow(i)
                            If sourceFieldName = "" Then
                                If i = currentRow.Length - 1 Then
                                    Continue For
                                Else
                                    sourceFieldName = "Column Number " & i + 1
                                End If

                            End If
                            If sourceFieldName = fieldNameToContainOldRecordIdNumbers Then 'this is just in case this field got backed up. It should be short lived
                                'only while QuNect Restore is running and doing a restore
                                Continue For
                            End If
                            If sourceFieldNames.ContainsKey(sourceFieldName) Then
                                Throw New System.Exception("Duplicate field label '" & sourceFieldName & "' in CSV file: " & lblFile.Text)
                            End If
                            If Not fieldTypes Is Nothing AndAlso fieldTypes.Length = currentRow.Length - 1 Then
                                sourceLabelToFieldType.Add(sourceFieldName, fieldTypes(i))
                            ElseIf Not clist Is Nothing AndAlso fieldNodes.ContainsKey(clist(i)) Then
                                sourceLabelToFieldType.Add(sourceFieldName, fieldNodes(clist(i)).type)
                            End If
                            dgMapping.Rows.Add(New String() {sourceFieldName})
                            sourceFieldNames.Add(sourceFieldName, i)
                            DirectCast(dgCriteria.Columns(filter.source), System.Windows.Forms.DataGridViewComboBoxColumn).Items.Add(sourceFieldName)
                            guessDestination(clist, sourceFieldName, j)
                            j += 1
                        Next

                    Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                        PopUpMsgBox("First line Of file Is malformed, cannot display field names.")

                    End Try
                End If
            End Using
            If Not bulkRestore Then

                Me.Cursor = Cursors.Default
            End If

        Catch ex As Exception
            PopUpMsgBox(ex.Message)
        End Try

    End Function
    Sub guessDestination(clist As String(), sourceFieldName As String, sourceFieldOrdinal As Integer)
        If Not clist Is Nothing AndAlso clist.Length > sourceFieldOrdinal Then
            For Each field As KeyValuePair(Of String, fieldStruct) In fieldNodes
                If field.Key = clist(sourceFieldOrdinal) AndAlso field.Value.type <> "address" Then
                    DirectCast(dgMapping.Rows(sourceFieldOrdinal).Cells(mapping.destination), System.Windows.Forms.DataGridViewComboBoxCell).Value = field.Value.label
                    Exit For
                End If
            Next
        Else
            For Each field As KeyValuePair(Of String, fieldStruct) In fieldNodes
                If field.Value.label = sourceFieldName AndAlso field.Value.type <> "address" Then
                    DirectCast(dgMapping.Rows(sourceFieldOrdinal).Cells(mapping.destination), System.Windows.Forms.DataGridViewComboBoxCell).Value = sourceFieldName
                    Exit For
                End If
            Next
        End If
    End Sub
    Private Sub ckbDetectProxy_CheckStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ckbDetectProxy.CheckStateChanged
        If ckbDetectProxy.Checked Then
            SaveSetting(AppName, "Credentials", "detectproxysettings", "1")
        Else
            SaveSetting(AppName, "Credentials", "detectproxysettings", "0")
        End If
    End Sub
    Private Sub chkBxHeaders_CheckStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkBxHeaders.CheckStateChanged
        If chkBxHeaders.Checked Then
            SaveSetting(AppName, "config", "firstrowhasheaders", "1")
        Else
            SaveSetting(AppName, "config", "firstrowhasheaders", "0")
        End If
    End Sub

    Function tableNameToDBID(tableName As String) As String
        Dim getDBIDfromdbName As New Regex("([a-z0-9~]+)$", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
        Dim dbidMatch As Match = getDBIDfromdbName.Match(tableName)
        tableNameToDBID = dbidMatch.Value
    End Function
    Private Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click
        Import()
    End Sub
    Private Sub Import()
        If cmbBulkorSingle.SelectedIndex = BulkOrSingle.TwoOrMore Then
            Me.Cursor = Cursors.WaitCursor
            btnSource.Enabled = False
            cmbBulkorSingle.Enabled = False
            Application.DoEvents()
            Dim strFolder As String = lblFile.Text
            Dim dbidToKeyFID As New Dictionary(Of String, String)
            Try
                Dim tableCounter As Integer = 0
                'import a whole directory
                Dim di As New DirectoryInfo(lblFile.Text)
                ' Get a reference to each file in that directory.
                Dim fiArr As FileInfo() = di.GetFiles("*.csv", SearchOption.AllDirectories)
                ' Display the names of the files.
                Dim fri As FileInfo
                Dim restoredDBIDs As New ArrayList
                For Each fri In fiArr
                    Me.Cursor = Cursors.WaitCursor
                    Application.DoEvents()
                    Dim tableName As String = Regex.Replace(fri.FullName, "\.csv$", "")
                    If Not System.IO.File.Exists(tableName & ".fids") Then
                        Continue For
                    End If
                    tableName = Regex.Replace(fri.Name, "\.csv$", "")
                    Dim dbid As String = tableNameToDBID(tableName)
                    restoredDBIDs.Add(dbid)
                    lblFile.Text = fri.FullName
                    lblTable.Text = tableName
                    Dim keyfid As String = listFieldsAndReturnKeyFID(tableName, True)
                    dbidToKeyFID.Add(dbid, keyfid)
                    Try
                        restoreTable(False, False, True)
                    Catch ex As Exception
                        PopUpMsgBox("Restoring " & tableName & " from " & fri.FullName & " failed because: " & ex.Message, MsgBoxStyle.OkOnly, AppName)
                    End Try

                    tableCounter += 1
                Next fri
                'now we have to go through all the dbids that we restored to see if there are any relationships between these tables
                'look for mastag fields
                Dim connectionString As String = getConnectionString(False, False)
                Using connection As OdbcConnection = getquNectConn(connectionString)
                    If connection Is Nothing Then Exit Sub
                    For Each dbid As String In restoredDBIDs
                        Application.DoEvents()
                        Dim strSQL As String = "select fid, mastag from """ & dbid & "~fields"" where mastag <> ''"
                        Using quNectCmd As OdbcCommand = New OdbcCommand(strSQL, connection)
                            Dim dr As OdbcDataReader = quNectCmd.ExecuteReader()
                            If Not dr.HasRows Then
                                quNectCmd.Dispose()
                                Continue For
                            End If
                            Dim aliasToDBID As New Dictionary(Of String, String)
                            Using quNectAliasesCmd As OdbcCommand = New OdbcCommand("select * from """ & dbid & "~aliases""", connection)
                                Dim drAliases As OdbcDataReader = quNectAliasesCmd.ExecuteReader()
                                While (drAliases.Read())
                                    aliasToDBID.Add(drAliases.GetString(0), drAliases.GetString(1))
                                End While
                                quNectAliasesCmd.Dispose()
                            End Using
                            While (dr.Read())
                                Application.DoEvents()
                                Dim referenceFID As String = dr.GetString(0)
                                Dim mastag As String = dr.GetString(1).ToLower()
                                If Not aliasToDBID.ContainsKey(mastag) Then Continue While
                                Dim parentDBID As String = aliasToDBID(mastag)
                                If Not dbidToKeyFID.ContainsKey(parentDBID) OrElse dbidToKeyFID(parentDBID) <> "3" Then
                                    Continue While
                                End If
                                Dim strReMapSQL As String = "INSERT INTO """ & dbid & """ (""" & dbid & """.fid" & dbidToKeyFID(dbid) & ", """ & dbid & """.fid" & referenceFID & ") SELECT """ & dbid & """.fid" & dbidToKeyFID(dbid) & ", " & parentDBID & ".fid3 FROM " & parentDBID & " INNER JOIN """ & dbid & """  ON " & parentDBID & ".""" & fieldNameToContainOldRecordIdNumbers & """ = """ & dbid & """.fid" & referenceFID & " WHERE " & parentDBID & ".""" & fieldNameToContainOldRecordIdNumbers & """ IS NOT NULL"
                                Using reMapConnection As OdbcConnection = getquNectConn(connectionString)
                                    If connection Is Nothing Then Exit Sub
                                    Using quNectReMapCmd As OdbcCommand = New OdbcCommand(strReMapSQL, reMapConnection)
                                        quNectReMapCmd.ExecuteNonQuery()
                                        quNectReMapCmd.Dispose()
                                    End Using
                                    reMapConnection.Close()
                                End Using
                            End While
                            quNectCmd.Dispose()
                        End Using
                    Next
                    connection.Close()

                End Using
                If lblDebug.Visible = False Then
                    For Each dbid As String In restoredDBIDs
                        Application.DoEvents()
                        If dbidToKeyFID(dbid) <> "3" Then
                            Continue For
                        End If
                        Using updateConnection As OdbcConnection = getquNectConn(connectionString)
                            Dim strSQL = "UPDATE " & dbid & " SET """ & fieldNameToContainOldRecordIdNumbers & """ = NULL"
                            Using quNectUpdateCmd As OdbcCommand = New OdbcCommand(strSQL, updateConnection)
                                quNectUpdateCmd.ExecuteNonQuery()
                                quNectUpdateCmd.Dispose()
                            End Using
                            updateConnection.Close()
                        End Using
                    Next
                End If
                Me.Cursor = Cursors.Default
                PopUpMsgBox("Restored " & tableCounter & " tables.", MsgBoxStyle.OkOnly, AppName)
            Catch ex As Exception
                Me.Cursor = Cursors.Default
                PopUpMsgBox("Could not import." & ex.Message, MsgBoxStyle.OkOnly, AppName)
            Finally
                Me.Cursor = Cursors.Default
                lblFile.Text = strFolder
                btnSource.Enabled = True
                cmbBulkorSingle.Enabled = True

            End Try

        ElseIf restoreTable(True, False, False) Then
            restoreTable(False, False, False)
        End If
    End Sub
    Private Sub PopUpMsgBox(msg As String)
        If cmdLineArgs.Length <= 1 Then
            MsgBox(msg)
        End If
    End Sub
    Private Sub PopUpMsgBox(msg As String, Style As MsgBoxStyle, Title As String)
        If cmdLineArgs.Length <= 1 Then
            MsgBox(msg, Style, Title)
        End If
    End Sub

    Private Sub dgCriteria_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles dgCriteria.CellEnter
        dtPicker.Visible = False
        If e.ColumnIndex <> filter.criteria Then
            dtPicker.Visible = False
            Exit Sub
        End If
        Dim label As String = sender.rows.item(e.RowIndex).cells.item(filter.source).value
        If label Is Nothing Then Exit Sub
        If sourceLabelToFieldType.ContainsKey(label) Then
            Dim field_type As String = sourceLabelToFieldType(label)
            If field_type = "date" Or field_type = "timestamp" Then
                'launch the date picker
                dtPicker.Top = ((DirectCast(e, DataGridViewCellEventArgs).RowIndex + 1) * dgCriteria.Rows(0).Height) + dgCriteria.Top
                dtPicker.Left = 43 + dgCriteria.Columns(filter.source).Width + dgCriteria.Columns(filter.criteria).Width + dgCriteria.Left
                dtPicker.Visible = True
                dtPicker.Tag = e.RowIndex
                dtPicker.BringToFront()
            End If
        End If
    End Sub



    Private Sub dtPicker_ValueChanged(sender As Object, e As EventArgs) Handles dtPicker.ValueChanged
        dgCriteria.Rows.Item(dtPicker.Tag).Cells.Item(filter.criteria).Value = dtPicker.Value
    End Sub

    Private Sub dgCriteria_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgCriteria.CellValueChanged
        If e.RowIndex < 0 Then Exit Sub
        If Not dgCriteria.Rows(e.RowIndex).Cells(filter.source).Value Is Nothing Then
            Select Case sourceLabelToFieldType(dgCriteria.Rows(e.RowIndex).Cells(filter.source).Value)
                Case "duration"
                    With dgCriteria.Rows(e.RowIndex).Cells(filter.criteria)
                        .ToolTipText = "Enter durations as Days.HH:MM:SS"
                    End With
                Case "checkbox"
                    With dgCriteria.Rows(e.RowIndex).Cells(filter.criteria)
                        .ToolTipText = "yes, 1, true means checked otherwise unchecked"
                    End With
            End Select
        End If
    End Sub

    Private Sub dgCriteria_CellMouseEnter(sender As Object, e As DataGridViewCellEventArgs) Handles dgCriteria.CellMouseEnter
        If e.RowIndex < 0 Then Exit Sub
        If Not dgCriteria.Rows(e.RowIndex).Cells(filter.source).Value Is Nothing AndAlso sourceLabelToFieldType.ContainsKey(dgCriteria.Rows(e.RowIndex).Cells(filter.source).Value) Then
            Select Case sourceLabelToFieldType(dgCriteria.Rows(e.RowIndex).Cells(filter.source).Value)
                Case "duration"
                    With dgCriteria.Rows(e.RowIndex).Cells(filter.criteria)
                        .ToolTipText = "Enter durations as Days.HH:MM:SS"
                    End With
                Case "checkbox"
                    With dgCriteria.Rows(e.RowIndex).Cells(filter.criteria)
                        .ToolTipText = "yes, 1, true means checked otherwise unchecked"
                    End With
            End Select
        End If
    End Sub

    Private Sub btnPreview_Click(sender As Object, e As EventArgs) Handles btnPreview.Click
        restoreTable(True, True, False)
    End Sub

    Private Sub dgMapping_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgMapping.CellMouseClick
        If e.RowIndex > 0 AndAlso e.ColumnIndex = mapping.source Then
            If dgMapping.Rows(e.RowIndex).Cells(mapping.destination).Value = "" Then
                guessDestination(clist, dgMapping.Rows(e.RowIndex).Cells(mapping.source).Value, e.RowIndex)
            Else
                dgMapping.Rows(e.RowIndex).Cells(mapping.destination).Value = ""
            End If

        End If
    End Sub


    Private Sub btnAppToken_Click(sender As Object, e As EventArgs) Handles btnAppToken.Click
        Process.Start("https://qunect.com/flash/AppToken.html")
    End Sub

    Private Sub btnUserToken_Click(sender As Object, e As EventArgs) Handles btnUserToken.Click
        Process.Start("https://qunect.com/flash/UserToken.html")
    End Sub

    Private Sub frmRestore_Click(sender As Object, e As MouseEventArgs) Handles Me.MouseClick
        If e.Button = MouseButtons.Right And (Control.ModifierKeys = (Keys.Shift Or Keys.Control)) Then
            lblDebug.Visible = True
        End If
    End Sub

    Private Sub lblDebug_MouseClick(sender As Object, e As MouseEventArgs) Handles lblDebug.MouseClick
        lblDebug.Visible = False
    End Sub

End Class


