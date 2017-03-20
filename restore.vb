
Imports System.Xml
Imports System.Net
Imports System.IO
Imports System.Text
Imports System.Data.Odbc
Imports System.Text.RegularExpressions
Imports System.Threading
Imports Microsoft.VisualBasic.FileIO
Imports System.Web


Public Class frmRestore
    Private Structure fieldStruct
        Public fid As String
        Public label As String
        Public type As String
        Public parentFieldID As String
        Public unique As Boolean
        Public required As Boolean
        Public base_type As String
        Public decimal_places As Integer
    End Structure
    Private qdb As QuickBaseClient
    Private Const AppName = "QuNectRestore"
    Private fieldNodes As Dictionary(Of String, fieldStruct)
    Private schemaXML As XmlDocument
    Private clist As String()
    Private fieldTypes As String()
    Private rids As HashSet(Of String)
    Private keyfid As String
    Private uniqueFieldValues As Hashtable
    Private sourceLabelToFieldType As Dictionary(Of String, String)
    Private sourceFieldNames As New Dictionary(Of String, Integer)
    Private destinationLabelsToFids As Dictionary(Of String, String)
    Private isBooleanTrue As Regex = New Regex("y|tr|c|[1-9]", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
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
    Private Function restoreTable(checkForErrorsOnly As Boolean, previewOnly As Boolean) As Boolean
        restoreTable = True

        Dim dsPreview As New DataSet
        Dim dtPreview As New DataTable
        Dim drPreview As DataRow = Nothing

        Dim dbid As String = Regex.Replace(lblTable.Text, "^.* ", "")
        'need to create the SQL statement
        Dim strSQL As String = "INSERT INTO " & dbid & " ("
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
        For i As Integer = 0 To dgMapping.Rows.Count - 1
            Dim destComboBoxCell As DataGridViewComboBoxCell = DirectCast(dgMapping.Rows(i).Cells(mapping.destination), System.Windows.Forms.DataGridViewComboBoxCell)
            If destComboBoxCell.Value Is Nothing Then Continue For
            Dim destDDIndex = destComboBoxCell.Items.IndexOf(destComboBoxCell.Value)

            If destDDIndex > 0 Then
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
                    MsgBox("You cannot import two different columns into the same field: " & destComboBoxCell.Value, MsgBoxStyle.OkOnly, AppName)
                    restoreTable = False
                    Exit Function
                End If
                fidsForImport.Add(fid)
                If fid = "3" Then ridIsMapped = True
                fids.Add(fid)
                vals.Add("?")
                importThese.Add(i)
                fieldLabels.Add(label)
            End If
        Next
        'now we need to look for unique fields
        Dim restrictions(1) As String
        restrictions(0) = dbid
        Dim connectionString As String = "Driver={QuNect ODBC For QuickBase};uid=" & txtUsername.Text & ";pwd=" & txtPassword.Text & ";QUICKBASESERVER=" & txtServer.Text & ";USEFIDS=1;APPTOKEN=" & txtAppToken.Text
        Using connection As OdbcConnection = getquNectConn(connectionString)
            If connection Is Nothing Then Exit Function
            'also need to pull all the fields marked required
            If checkForErrorsOnly Then
            frmErr.rdbCancel.Checked = True
            uniqueFieldValues = New Hashtable

                For Each field As KeyValuePair(Of String, fieldStruct) In fieldNodes
                    Dim fid As String = field.Value.fid

                    If field.Value.unique Then
                        If Not fidsForImport.Contains(fid) Then
                            requiredFieldsErrors &= "You must import values into the unique field: " & field.Value.label
                            restoreTable = False
                        ElseIf fid <> keyfid Then
                            'we also need to keep a list of the unique fields that we're importing into to check and make sure all imported values are unique
                            uniqueFidsForImport.Add(fid)
                            'do we need this hashset to check the imported values? We would need to check to see if we were updating an existing record or creating a new record
                            'if we have a 
                            If Not ridIsMapped Then
                                uniqueFieldValues.Add(fid, qdb.getHashSetofFieldValues(dbid, fid))
                            End If
                        End If
                    End If
                    If field.Value.required Then
                        If Not fidsForImport.Contains(fid) Then
                            requiredFieldsErrors &= "You must import values into the required field: " & field.Value.label
                            restoreTable = False
                        Else
                            'we also need to keep a list of the required fields that we're importing into
                            requiredFidsForImport.Add(fid)
                        End If
                    End If
                Next
            End If
        If ridIsMapped And checkForErrorsOnly Then
            'need to pull in all the rids from QuickBase
            rids = qdb.getHashSetofFieldValues(dbid, "3")
        Else
            rids.Clear()
        End If
            strSQL &= "fid" & String.Join(",fid", fids.ToArray)
            strSQL &= ") VALUES ("
            strSQL &= String.Join(",", vals.ToArray)
            strSQL &= ")"
            Using command As OdbcCommand = New OdbcCommand(strSQL, connection)
                For i As Integer = 0 To fids.Count - 1
                    Dim qdbType As OdbcType = getODBCTypeFromQuickBaseFieldNode(fieldNodes(fids(i)))
                    command.Parameters.Add("@fid" & fids(i), qdbType) 'we need to set the type here according to the field schema
                Next
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
                            For i As Integer = 0 To importThese.Count - 1
                                Dim val As String = currentRow(importThese(i))
                                If previewOnly Then
                                    drPreview(i) = val
                                End If
                                If checkForErrorsOnly Then
                                    If fids(i) = "3" AndAlso Not rids.Contains(val) Then
                                        'here we're going to have a problem because this record no longer exists in QuickBase
                                        'we could update all the child tables with the newly minted Record ID#s
                                        'but we wouid have to do one record at a time to accomplish this
                                        missingRIDs &= vbCrLf & "line " & fileLineCounter + 1 & " has Record ID# " & val & " which no longer exists"
                                        If missingRIDs.Length > 1000 Then
                                            missingRIDs &= vbCrLf & "There may be additional errors beyond the ones above."
                                            Exit While
                                        End If
                                    End If
                                    If requiredFidsForImport.Contains(fids(i)) Then
                                        If val = "" Then
                                            requiredFieldsErrors &= vbCrLf & "line " & fileLineCounter + 1 & " has a blank value mapped To required field " & fieldLabels(i)
                                            If requiredFieldsErrors.Length > 1000 Then
                                                requiredFieldsErrors &= vbCrLf & "There may be additional errors beyond the ones above."
                                                Exit While
                                            End If
                                        End If
                                    End If
                                    If uniqueFidsForImport.Contains(fids(i)) Then
                                        'here we need to see if we have a duplicate
                                        If Not uniqueFieldValues.Contains(fids(i)) Then
                                            uniqueFieldValues.Add(fids(i), New HashSet(Of String))
                                        End If
                                        If uniqueFieldValues(fids(i)).contains(val) Then
                                            uniqueFieldErrors &= vbCrLf & "line " & fileLineCounter + 1 & " has a duplicate value For unique field " & fieldLabels(i)
                                        Else
                                            uniqueFieldValues(fids(i)).add(val)
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
                            If previewOnly Then
                                dtPreview.Rows.Add(drPreview)
                            End If
                            If Not checkForErrorsOnly Then command.ExecuteNonQuery()

                        Catch ex As Exception
                            If checkForErrorsOnly Then
                                improperlyFormattedLines &= vbCrLf & "line " & fileLineCounter + 1 & " Is Not a properly formatted CSV line."
                                If improperlyFormattedLines.Length > 1000 Then
                                    improperlyFormattedLines &= vbCrLf & "There may be additional errors beyond the ones above."
                                    Exit While
                                End If
                            End If
                        End Try
                        lineCounter += 1
                        pb.Value = fileLineCounter Mod 1000
                    End While
                    If Not checkForErrorsOnly Then
                        transaction.Commit()
                        MsgBox("Imported " & lineCounter & " records!", MsgBoxStyle.OkOnly, AppName)
                    Else
                        restoreTable = showErrors(previewOnly, conversionErrors, improperlyFormattedLines, uniqueFieldErrors, requiredFieldsErrors, missingRIDs)
                        pb.Visible = False
                    End If
                    If previewOnly And restoreTable Then
                        dsPreview.Tables.Add(dtPreview)
                        frmPreview.dgPreview.DataSource = dsPreview.Tables(0)
                        frmPreview.ShowDialog()
                    End If
                Catch ex As Exception
                    MsgBox("Could Not import because " & ex.Message)
                    restoreTable = False
                Finally
                    lblProgress.Text = ""
                    pb.Visible = False
                End Try
            End Using
        End Using
    End Function
    Private Function getquNectConn(connectionString As String) As OdbcConnection

        Dim quNectConn As OdbcConnection = New OdbcConnection(connectionString)
        Try
            quNectConn.Open()
        Catch excpt As Exception
            Me.Cursor = Cursors.Default
            If excpt.Message.StartsWith("Error [IM003]") Or excpt.Message.Contains("Data source name Not found") Then
                MsgBox("Please install QuNect ODBC For QuickBase from http://qunect.com/download/QuNect.exe and try again.")
            Else
                MsgBox(excpt.Message.Substring(13))
            End If
            Return Nothing
            Exit Function
        End Try

        Dim ver As String = quNectConn.ServerVersion
        Dim m As Match = Regex.Match(ver, "\d+\.(\d+)\.(\d+)\.(\d+)")
        qdbVer.year = CInt(m.Groups(1).Value)
        qdbVer.major = CInt(m.Groups(2).Value)
        qdbVer.minor = CInt(m.Groups(3).Value)
        If (qdbVer.major < 6) Or (qdbVer.major = 6 And qdbVer.minor < 84) Then
            MsgBox("You are running the " & ver & " version of QuNect ODBC for QuickBase. Please install the latest version from http://qunect.com/download/QuNect.exe")
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
                    command.Parameters("@fid" & fid).Value = Convert.ToInt32(val)
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

    Private Sub txtUsername_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUsername.TextChanged
        SaveSetting(AppName, "Credentials", "username", txtUsername.Text)
    End Sub

    Private Sub txtPassword_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPassword.TextChanged
        SaveSetting(AppName, "Credentials", "password", txtPassword.Text)
    End Sub
    Private Sub btnSource_Click(sender As Object, e As EventArgs) Handles btnSource.Click
        OpenSourceFile.ShowDialog()
    End Sub

    Private Sub OpenSourceFile_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenSourceFile.FileOk
        lblFile.Text = OpenSourceFile.FileName.ToString()
        hideButtons()
    End Sub
    Private Sub hideButtons()
        btnPreview.Visible = False
        btnImport.Visible = False
        dgCriteria.Visible = False
        dgMapping.Visible = False
    End Sub
    Private Sub btnListTables_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnListTables.Click
        listTables()
    End Sub
    Private Sub listTables()
        Me.Cursor = Cursors.WaitCursor

        qdb.setServer(txtServer.Text, True)
        qdb.setAppToken(txtAppToken.Text)
        qdb.Authenticate(txtUsername.Text, txtPassword.Text)
        Try
            qdb.GetGrantedDBs(True, True, False, New AsyncCallback(AddressOf listTablesCallback), New AsyncCallback(AddressOf timeoutCallback))
        Catch ex As Exception
            Me.Cursor = Cursors.Default
            MsgBox(ex.Message)
            Exit Sub
        End Try
    End Sub
    Sub timeoutCallback(ByVal result As System.IAsyncResult)
        Me.Cursor = Cursors.Default
        MsgBox("Operation timed out. Please Try again.")
    End Sub
    Sub listTablesCallback(ByVal result As System.IAsyncResult)
        Dim tableXML As XmlDocument
        Dim tableNodes As XmlNodeList

        Dim requestState As QuickBaseClient.WebRequestState = CType(result.AsyncState, QuickBaseClient.WebRequestState)
        tableXML = requestState.xmlDoc
        tableNodes = tableXML.SelectNodes("/*/databases/dbinfo")
        Dim defaultdbid As String = GetSetting(AppName, "restore", "dbid").ToLower

        frmTableChooser.tvAppsTables.BeginUpdate()
        frmTableChooser.tvAppsTables.Nodes.Clear()
        Dim dbName As String
        Dim applicationName As String = ""
        Dim prevAppName As String = ""
        Dim dbid As String
        pb.Value = 0
        pb.Visible = True
        pb.Maximum = tableNodes.Count
        For i = 0 To tableNodes.Count - 1
            pb.Value = i
            Application.DoEvents()
            dbName = tableNodes(i).SelectSingleNode("dbname").InnerText
            applicationName = dbName.Split(":")(0)
            dbid = tableNodes(i).SelectSingleNode("dbid").InnerText

            If applicationName <> prevAppName Then
                frmTableChooser.tvAppsTables.Nodes.Add(applicationName)
                prevAppName = applicationName
            End If
            Dim tableName As String = dbName
            If dbName.Length > applicationName.Length Then
                tableName = dbName.Substring(applicationName.Length + 1)
            End If
            frmTableChooser.tvAppsTables.Nodes(frmTableChooser.tvAppsTables.Nodes.Count - 1).Nodes.Add(tableName & " " & dbid)
            If defaultdbid = dbid Then
                lblTable.Text = frmTableChooser.tvAppsTables.Nodes(frmTableChooser.tvAppsTables.Nodes.Count - 1).Nodes(frmTableChooser.tvAppsTables.Nodes(frmTableChooser.tvAppsTables.Nodes.Count - 1).Nodes.Count - 1).FullPath()
            End If
        Next
        pb.Visible = False
        frmTableChooser.tvAppsTables.EndUpdate()
        pb.Value = 0
        btnImport.Visible = True
        lblTable.Visible = True
        frmTableChooser.Show()
        Me.Cursor = Cursors.Default
    End Sub
    Sub listFields(dbid As String)
        sourceLabelToFieldType = New Dictionary(Of String, String)
        destinationLabelsToFids = New Dictionary(Of String, String)
        Dim fidToLabel As New Dictionary(Of String, String)
        fieldNodes = New Dictionary(Of String, fieldStruct)
        sourceFieldNames.Clear()
        Dim connectionString As String = "Driver={QuNect ODBC for QuickBase};uid=" & txtUsername.Text & ";pwd=" & txtPassword.Text & ";QUICKBASESERVER=" & txtServer.Text & ";APPTOKEN=" & txtAppToken.Text
        Try
            Dim currentRow As String()
            Using connection As OdbcConnection = getquNectConn(connectionString)
                If connection Is Nothing Then Exit Sub
                Dim strSQL As String = "SELECT label, fid, field_type, parentFieldID, ""unique"", required, ""key"", base_type, decimal_places  FROM """ & dbid & "~fields"" WHERE (mode = '' and role = '') or fid = '3'"

                Dim quNectCmd As OdbcCommand = New OdbcCommand(strSQL, connection)
                Dim dr As OdbcDataReader
                Try
                    dr = quNectCmd.ExecuteReader()
                Catch excpt As Exception
                    quNectCmd.Dispose()
                    Exit Sub
                End Try
                If Not dr.HasRows Then
                    Exit Sub
                End If
                Dim fidReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(Regex.Replace(lblFile.Text, "csv$", "fids"))
                fidReader.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited
                fidReader.Delimiters = New String() {"."}
                clist = fidReader.ReadFields()
                If Not fidReader.EndOfData Then
                    fieldTypes = fidReader.ReadFields()
                Else
                    fieldTypes = Nothing
                End If
                'here we need to open the csv and get the field names
                Dim csvReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(lblFile.Text)
                csvReader.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited
                csvReader.Delimiters = New String() {","}
                csvReader.HasFieldsEnclosedInQuotes = True
                'Loop through all of the fields in the schema.
                DirectCast(dgMapping.Columns(mapping.destination), System.Windows.Forms.DataGridViewComboBoxColumn).Items.Clear()
                DirectCast(dgCriteria.Columns(filter.source), System.Windows.Forms.DataGridViewComboBoxColumn).Items.Clear()
                DirectCast(dgMapping.Columns(mapping.destination), System.Windows.Forms.DataGridViewComboBoxColumn).Items.Add("Do not import")
                While (dr.Read())
                    Dim field As New fieldStruct
                    field.label = dr.GetString(0)
                    field.fid = dr.GetString(1)
                    field.type = dr.GetString(2)
                    field.parentFieldID = dr.GetString(3)
                    field.unique = dr.GetBoolean(4)
                    field.required = dr.GetBoolean(5)
                    fidToLabel.Add(field.fid, field.label)
                    If field.parentFieldID <> "" Then
                        field.label = fidToLabel(field.parentFieldID) & ": " & field.label
                    End If
                    If dr.GetBoolean(6) Then
                        keyfid = field.fid
                    End If
                    field.base_type = dr.GetString(7)
                    field.decimal_places = 0
                    If Not IsDBNull(dr(8)) Then
                        field.decimal_places = dr.GetInt32(8)
                    End If
                    fieldNodes.Add(field.fid, field)
                    destinationLabelsToFids.Add(field.label, field.fid)
                    DirectCast(dgMapping.Columns(mapping.destination), System.Windows.Forms.DataGridViewComboBoxColumn).Items.Add(field.label)
                End While
                quNectCmd.Dispose()
                If Not csvReader.EndOfData Then
                    Try
                        currentRow = csvReader.ReadFields()
                        dgMapping.Rows.Clear()
                        For i As Integer = 0 To currentRow.Length - 1
                            Dim sourceFieldName As String = currentRow(i)
                            If sourceFieldName = "" Then
                                Exit For
                            End If
                            If fieldTypes.Length = currentRow.Length - 1 Then
                                sourceLabelToFieldType.Add(sourceFieldName, fieldTypes(i))
                            ElseIf fieldNodes.ContainsKey(clist(i)) Then
                                sourceLabelToFieldType.Add(sourceFieldName, fieldNodes(clist(i)).type)
                            End If
                            If sourceFieldName = "" Then Continue For
                            Dim j As Integer = dgMapping.Rows.Add(New String() {sourceFieldName})
                            sourceFieldNames.Add(sourceFieldName, i)
                            DirectCast(dgCriteria.Columns(filter.source), System.Windows.Forms.DataGridViewComboBoxColumn).Items.Add(sourceFieldName)
                            'DirectCast(dgCriteria.Rows(j).Cells(filter.booleanOperator), System.Windows.Forms.DataGridViewComboBoxCell).Value = "has any value"
                            'now we have to see if we can match this to a field in the table
                            'DirectCast(dgMapping.Rows(j).Cells(mapping.destination), System.Windows.Forms.DataGridViewComboBoxCell).Value = "Do not import"
                            If clist.Length > i Then
                                For Each field As KeyValuePair(Of String, fieldStruct) In fieldNodes
                                    If field.Key = clist(i) AndAlso field.Value.type <> "address" Then
                                        DirectCast(dgMapping.Rows(j).Cells(mapping.destination), System.Windows.Forms.DataGridViewComboBoxCell).Value = field.Value.label
                                        Exit For
                                    End If
                                Next
                            Else
                                For Each field As KeyValuePair(Of String, fieldStruct) In fieldNodes
                                    If field.Value.label = sourceFieldName AndAlso field.Value.type <> "address" Then
                                        DirectCast(dgMapping.Rows(j).Cells(mapping.destination), System.Windows.Forms.DataGridViewComboBoxCell).Value = sourceFieldName
                                        Exit For
                                    End If
                                Next
                            End If
                        Next

                    Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                        MsgBox("First line of file is malformed, cannot display field names.")

                    End Try
                End If
            End Using

            btnImport.Visible = True
            btnPreview.Visible = True
            dgMapping.Visible = True
            dgCriteria.Visible = True
            Me.Cursor = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
    Private Sub txtServer_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtServer.TextChanged
        SaveSetting(AppName, "Credentials", "server", txtServer.Text)
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
    Private Sub restore_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.Text = "QuNect Restore 1.0.0.1"
        txtUsername.Text = GetSetting(AppName, "Credentials", "username")
        txtPassword.Text = GetSetting(AppName, "Credentials", "password")
        txtServer.Text = GetSetting(AppName, "Credentials", "server", "www.quickbase.com")
        txtAppToken.Text = GetSetting(AppName, "Credentials", "apptoken", "b2fr52jcykx3tnbwj8s74b8ed55b")
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

        qdb = New QuickBaseClient(txtUsername.Text, txtPassword.Text)
        qdb.proxyAuthenticateEx(True)

        Dim myBuildInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(Application.ExecutablePath)
        Dim ops As String() = New String() {"has any value", "equals", "does not equal", "greater than", "greater than or equal", "less than", "less than or equal", "contains", "does not contain", "starts with", "does not start with", "is null", "is not null"}
        For i As Integer = 0 To ops.Count - 1
            DirectCast(dgCriteria.Columns(filter.booleanOperator), System.Windows.Forms.DataGridViewComboBoxColumn).Items.Add(ops(i))
        Next
        dgMapping.Visible = False
        dgCriteria.Visible = False
    End Sub

    Private Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click
        If restoreTable(True, False) Then
            restoreTable(False, False)
        End If
    End Sub

    Private Sub btnListFields_Click(sender As Object, e As EventArgs) Handles btnListFields.Click
        If lblFile.Text = "" Then
            MsgBox("Please choose a file to import.", MsgBoxStyle.OkOnly, AppName)
            Me.Cursor = Cursors.Default
            Exit Sub
        ElseIf lblTable.Text = "" Then
            MsgBox("Please choose a table to import.", MsgBoxStyle.OkOnly, AppName)
            Me.Cursor = Cursors.Default
            Exit Sub
        End If
        'qdb.setServer(txtServer.Text, True)
        'qdb.setAppToken(txtAppToken.Text)
        'qdb.Authenticate(txtUsername.Text, txtPassword.Text)
        'qdb.GetSchema(lblTable.Text, New AsyncCallback(AddressOf schemaCallback), New AsyncCallback(AddressOf timeoutCallback))
        listFields(lblTable.Text)
    End Sub


    Private Sub txtAppToken_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAppToken.TextChanged
        SaveSetting(AppName, "Credentials", "apptoken", txtAppToken.Text)
    End Sub
    Private Sub lblTable_TextChanged(sender As Object, e As EventArgs) Handles lblTable.TextChanged
        SaveSetting(AppName, "config", "table", lblTable.Text)
        btnImport.Visible = False
    End Sub

    Private Sub lblFile_TextChanged(sender As Object, e As EventArgs) Handles lblFile.TextChanged
        SaveSetting(AppName, "config", "file", lblFile.Text)
        btnImport.Visible = False
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
        restoreTable(True, True)
    End Sub
End Class


Public Class QuickBaseClient

    Private Password As String
    Private UserName As String
    Private strProxyPassword As String
    Private strProxyUsername As String
    Private ticket As String
    Private apptoken As String
    Private QDBHost As String = "www.quickbase.com"
    Private useHTTPS As Boolean = True
    Public GMTOffset As Single


    Public errorcode As Integer
    Public errortext As String
    Public errordetail As String
    Public httpContentLengthProgress As Integer
    Public httpContentLength As Integer

    Private Const OB32CHARACTERS As String = "abcdefghijkmnpqrstuvwxyz23456789"
    Private Const Map64 As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    Private Const MILLISECONDS_IN_A_DAY As Double = 86400000.0#
    Private Const DAYS_BETWEEN_JAVASCRIPT_AND_MICROSOFT_DATE_REFERENCES As Double = 25569.0#

    Private proxy As IWebProxy = Nothing
    Public Sub proxyAuthenticateEx(ByVal useDefaultProxy As Boolean, Optional ByVal proxyServerUrl As String = "", Optional ByVal port As Integer = 80, Optional ByVal strUsername As String = "", Optional ByVal strPassword As String = "")
        If (useDefaultProxy) Then
            proxy = WebRequest.GetSystemWebProxy()
            proxy.Credentials = CredentialCache.DefaultCredentials
        Else
            proxy = New WebProxy(proxyServerUrl, port)
            proxy.Credentials = New NetworkCredential(strUsername, strPassword)
        End If
    End Sub
    Function makeCSVCells(ByRef cells As ArrayList) As String
        Dim i As Integer
        Dim cell As String
        makeCSVCells = ""
        For i = 0 To cells.Count - 1
            If cells(i) Is Nothing Then
                cell = ""
            Else
                cell = cells(i).ToString()
            End If
            makeCSVCells = makeCSVCells & """" & cell.Replace("""", """""") & ""","
        Next
    End Function

    Function encode32(ByVal strDecimal As String) As String

        Dim ob32 As String = ""
        Dim intDecimal As Integer
        intDecimal = CInt(strDecimal)
        Dim remainder As Integer

        Do While (intDecimal > 0)
            remainder = intDecimal Mod 32
            ob32 = Mid(OB32CHARACTERS, CInt(remainder) + 1, 1) & ob32
            intDecimal = intDecimal \ 32
        Loop
        encode32 = ob32

    End Function
    Public Function getTextByFID(ByRef recordNode As XmlNode, ByRef fid As String) As String
        Dim cell As XmlNode = recordNode.SelectSingleNode("f[@id=" & fid & "]")
        If cell Is Nothing Then
            Err.Raise(vbObjectError + 5, "QuickBase.QuickBaseClient", "Could not find fid " & fid)
        End If
        getTextByFID = cell.InnerText
    End Function
    Public Function makeClist(ByVal fids As Hashtable) As String
        Dim period As String = ""
        makeClist = ""
        For Each fid As DictionaryEntry In fids
            makeClist = makeClist & period & fid.Value
            period = "."
        Next
    End Function
    Public Function makeClist(ByRef fids As ArrayList) As String
        Dim period As String = ""
        makeClist = ""
        Dim i As Integer
        For i = 0 To fids.Count - 1
            makeClist = makeClist & period & fids(i)
            period = "."
        Next
    End Function
    Public Function makeClist(ByRef fids() As String) As String
        Dim period As String = ""
        makeClist = ""
        Dim fid As String
        For Each fid In fids
            makeClist = makeClist & period & fid
            period = "."
        Next
    End Function

    Public Function Authenticate(ByVal strUsername As String, ByVal strPassword As String) As Integer
        UserName = strUsername
        Password = strPassword
        ticket = ""
        Authenticate = 0
    End Function
    Public Function proxyAuthenticate(ByVal strUsername As String, ByVal strPassword As String) As Integer
        strProxyUsername = strUsername
        strProxyPassword = strPassword
        proxyAuthenticate = 0
    End Function

    Public Sub GetGrantedDBs(ByVal withEmbeddedTables As Boolean, ByVal excludeParents As Boolean, ByVal adminOnly As Boolean, ByRef callback As AsyncCallback, ByRef timeoutCallback As AsyncCallback)
        Dim xmlQDBRequest As XmlDocument

        xmlQDBRequest = InitXMLRequest()
        If withEmbeddedTables Then
            addParameter(xmlQDBRequest, "withEmbeddedTables", "1")
        Else
            addParameter(xmlQDBRequest, "withEmbeddedTables", "0")
        End If
        If excludeParents Then
            addParameter(xmlQDBRequest, "excludeParents", "1")
        Else
            addParameter(xmlQDBRequest, "excludeParents", "0")
        End If
        If adminOnly Then
            addParameter(xmlQDBRequest, "adminOnly", "1")
        End If
        APIXMLPost("main", "API_GrantedDBs", xmlQDBRequest, useHTTPS, callback, timeoutCallback)
    End Sub
    Public Function getHashSetofFieldValues(dbid As String, fid As String) As HashSet(Of String)
        Dim CSVStream As System.IO.Stream = GenResultsTable(dbid, "{'0'.CT.''}", fid, "", "csv")
        Dim CSVReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(CSVStream)
        CSVReader.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited
        CSVReader.Delimiters = New String() {","}
        CSVReader.HasFieldsEnclosedInQuotes = True
        Dim currentRow As String()
        If Not CSVReader.EndOfData Then currentRow = CSVReader.ReadFields()
        getHashSetofFieldValues = New HashSet(Of String)
        While Not CSVReader.EndOfData
            currentRow = CSVReader.ReadFields()
            getHashSetofFieldValues.Add(currentRow(0))
        End While
    End Function

    Public Function GenResultsTable(ByVal dbid As String, ByVal query As String, ByVal clist As String, ByVal slist As String, ByVal options As String) As System.IO.Stream
        Dim querystring As String = initQueryString()

        If CStr(Val(query)) = query Then
            querystring &= "&qid=" & System.Web.HttpUtility.UrlEncode(query)
        Else
            querystring &= "&query=" & System.Web.HttpUtility.UrlEncode(query)
        End If
        querystring &= "&clist=" & System.Web.HttpUtility.UrlEncode(clist)
        querystring &= "&slist=" & System.Web.HttpUtility.UrlEncode(slist)
        querystring &= "&options=" & System.Web.HttpUtility.UrlEncode(options)
        GenResultsTable = APIHTML(dbid, "API_GenResultsTable", querystring)
    End Function
    Public Sub GetSchema(ByVal dbid As String, ByRef callback As AsyncCallback, ByRef timeoutCallback As AsyncCallback)
        Dim xmlQDBRequest As XmlDocument
        dbid = Regex.Replace(dbid, "^.* ", "")
        xmlQDBRequest = InitXMLRequest()
        APIXMLPost(dbid, "API_GetSchema", xmlQDBRequest, useHTTPS, callback, timeoutCallback)
    End Sub
    ' Stores web request for access during async processing
    Public Class WebRequestState
        ' Holds the request object
        Public request As WebRequest
        Public response As HttpWebResponse
        Public callback As AsyncCallback
        Public timeoutCallback As AsyncCallback
        Public xmlDoc As XmlDocument
        Public Sub New(ByVal newRequest As WebRequest, ByRef newCallback As AsyncCallback, ByRef newTimeoutCallback As AsyncCallback)
            request = newRequest
            callback = newCallback
            timeoutCallback = newTimeoutCallback
        End Sub
    End Class
    Sub APIXMLPost(ByVal dbid As String, ByVal action As String, ByRef xmlQDBRequest As XmlDocument, ByVal useHTTPS As Boolean, ByRef callback As AsyncCallback, ByRef timeoutCallback As AsyncCallback)

        Dim script As String
        Dim content As String
        Dim req As HttpWebRequest

        script = QDBHost & "/db/" & dbid & "?act=" & action
        If useHTTPS Then
            script = "https://" & script
        Else
            script = "http://" & script
        End If
        Application.DoEvents()
        content = "<?xml version=""1.0"" encoding=""ISO-8859-1""?>" & xmlQDBRequest.OuterXml
        req = CType(WebRequest.Create(script), HttpWebRequest)
        req.ContentType = "text/xml"
        req.Method = "POST"
        If (proxy IsNot Nothing) Then
            req.Proxy = proxy
            req.PreAuthenticate = True
        End If
        Dim byteRequestArray As Byte() = Encoding.UTF8.GetBytes(content)
        req.ContentLength = byteRequestArray.Length
        Dim reqStream As Stream = req.GetRequestStream()
        Application.DoEvents()
        reqStream.Write(byteRequestArray, 0, byteRequestArray.Length)
        Application.DoEvents()
        ' Create the state object used to access the web request
        Dim state As WebRequestState = New WebRequestState(req, callback, timeoutCallback)

        Dim result As IAsyncResult = req.BeginGetResponse(New AsyncCallback(AddressOf APIXMLPostCallback), state)
        ' Set timeout at 1 minute
        Dim timeout As Integer = 1000 * 60

        ' Register a timeout for the async request
        ThreadPool.RegisterWaitForSingleObject(result.AsyncWaitHandle, New WaitOrTimerCallback(AddressOf APIXMLPostTimeoutCallback), state, timeout, True)
        Application.DoEvents()
        reqStream.Close()
        Application.DoEvents()

    End Sub
    Private Sub APIXMLPostCallback(ByVal result As IAsyncResult)
        Dim xmlStream As Stream
        Dim xmlTxtReader As XmlTextReader
        Dim xmlDoc As XmlDocument
        Dim request As WebRequest

        Dim requestState As WebRequestState = CType(result.AsyncState, WebRequestState)
        Dim httpWebRequest As HttpWebRequest = requestState.request
        requestState.response = CType(httpWebRequest.EndGetResponse(result), HttpWebResponse)

        'create a new stream that can be placed into an XmlTextReader
        xmlStream = requestState.response.GetResponseStream()
        Application.DoEvents()
        xmlTxtReader = New XmlTextReader(xmlStream)
        Application.DoEvents()
        xmlTxtReader.XmlResolver = Nothing
        'create a new Xml document
        xmlDoc = New XmlDocument
        Application.DoEvents()
        xmlDoc.Load(xmlTxtReader)
        xmlStream.Close()
        On Error Resume Next
        errorcode = CInt(requestState.response.Headers("QUICKBASE-ERRCODE"))
        ticket = xmlDoc.DocumentElement.SelectSingleNode("/*/ticket").InnerText
        errortext = xmlDoc.DocumentElement.SelectSingleNode("/*/errtext").InnerText
        If xmlDoc.DocumentElement.SelectSingleNode("/*/errdetail") Is Nothing Then
            errordetail = xmlDoc.DocumentElement.SelectSingleNode("/*/errtext").InnerText
        Else
            errordetail = xmlDoc.DocumentElement.SelectSingleNode("/*/errdetail").InnerText
        End If
        On Error GoTo 0
        If errorcode <> 0 Then
            Err.Raise(vbObjectError + CInt(errorcode), "QuickBase.QuickBaseClient", errordetail)
        End If
        requestState.xmlDoc = xmlDoc
        requestState.callback.Invoke(result)
    End Sub
    Private Sub APIXMLPostTimeoutCallback(ByVal state As Object, ByVal timeOut As Boolean)
        If (timeOut) Then
            ' Abort the request
            CType(state, WebRequestState).request.Abort()
            state.timeoutcallback.Invoke()
        End If



    End Sub
    Public Function APIHTMLPost(ByVal dbid As String, ByVal action As String, ByRef xmlQDBRequest As XmlDocument, ByVal useHTTPS As Boolean) As String

        Dim script As String


        script = "/db/" & dbid & "?act=" & action
        APIHTMLPost = HTTPPost(QDBHost, useHTTPS, script, "text/xml", xmlQDBRequest.OuterXml, "")

    End Function
    Public Function APIHTML(ByVal dbid As String, ByVal action As String, querystring As String) As System.IO.Stream

        Dim script As String


        script = "/db/" & dbid & "?act=" & action & querystring
        APIHTML = HTTPPost(QDBHost, True, script)

    End Function
    Private Function HTTPPost(ByVal QDBHost As String, ByVal useHTTPS As Boolean, ByVal script As String, ByVal contentType As String, ByVal content As String, ByVal fileName As String) As String
        Dim url As String
        Dim Client As WebClient

        Client = New WebClient
        Client.Headers.Add("Content-Type", contentType)
        url = QDBHost & script
        If useHTTPS Then
            url = "https://" & url
        Else
            url = "http://" & url
        End If
        Dim byteRequestArray As Byte() = Encoding.UTF8.GetBytes(content)

        Dim byteResponseArray As Byte() = Client.UploadData(url, "POST", byteRequestArray)
        If fileName = "" Then
            HTTPPost = Encoding.UTF8.GetString(byteResponseArray)
        Else
            'check if write file exists 
            If File.Exists(path:=fileName) Then
                'delete file
                File.Delete(path:=fileName)
            End If

            'create a fileStream instance to pass to BinaryWriter object
            Dim fsWrite As FileStream
            fsWrite = New FileStream(path:=fileName,
                mode:=FileMode.CreateNew, access:=FileAccess.Write)

            'create binary writer instance
            Dim bWrite As BinaryWriter
            bWrite = New BinaryWriter(output:=fsWrite)
            'write bytes out 
            bWrite.Write(byteResponseArray, 0, byteResponseArray.Length)


            'close the writer 
            bWrite.Close()
            HTTPPost = fileName
        End If

    End Function
    Private Function HTTPPost(ByVal QDBHost As String, ByVal useHTTPS As Boolean, ByVal script As String) As System.IO.Stream
        Dim url As String
        Dim Client As WebClient

        Client = New WebClient
        url = QDBHost & script
        If useHTTPS Then
            url = "https://" & url
        Else
            url = "http://" & url
        End If

        HTTPPost = Client.OpenRead(url)
    End Function
    Public Function setAppToken(ByVal aapptoken As String) As String
        apptoken = aapptoken
        setAppToken = apptoken
    End Function
    Public Function getServer() As String
        getServer = QDBHost
    End Function
    Private Function initQueryString() As String
        initQueryString = "&username=" & System.Web.HttpUtility.UrlEncode(UserName) & "&password=" & System.Web.HttpUtility.UrlEncode(Password) & "&apptoken=" & System.Web.HttpUtility.UrlEncode(apptoken)
    End Function


    Public Function InitXMLRequest() As XmlDocument
        Dim xmlQDBRequest As New XmlDocument
        Dim Root As XmlElement

        Root = xmlQDBRequest.CreateElement("qdbapi")
        xmlQDBRequest.AppendChild(Root)
        If Len(ticket) <> 0 Then
            addParameter(xmlQDBRequest, "ticket", ticket)
        ElseIf UserName <> "" Then
            addParameter(xmlQDBRequest, "username", UserName)
            addParameter(xmlQDBRequest, "password", Password)
        End If
        If Len(apptoken) <> 0 Then
            addParameter(xmlQDBRequest, "apptoken", apptoken)
        End If
        InitXMLRequest = xmlQDBRequest
    End Function

    Public Sub addParameter(ByRef xmlQDBRequest As XmlDocument, ByVal Name As String, ByVal Value As String)
        Dim Root As XmlElement
        Dim ElementNode As XmlNode
        Dim TextNode As XmlNode

        Root = xmlQDBRequest.DocumentElement
        ElementNode = xmlQDBRequest.CreateNode(XmlNodeType.Element, Name, "")
        TextNode = xmlQDBRequest.CreateNode(XmlNodeType.Text, "", "")
        TextNode.InnerText = Value
        ElementNode.AppendChild(TextNode)
        Root.AppendChild(ElementNode)
        Root = Nothing
        ElementNode = Nothing
        TextNode = Nothing
    End Sub

    Public Sub addParameterWithAttribute(ByRef xmlQDBRequest As XmlDocument, ByVal Name As String, ByVal AttributeName As String, ByVal AttributeValue As String, ByVal Value As String)
        Dim Root As XmlElement
        Dim ElementNode As XmlNode
        Dim TextNode As XmlNode
        Dim Attribute As XmlAttribute

        Root = xmlQDBRequest.DocumentElement
        ElementNode = xmlQDBRequest.CreateNode(XmlNodeType.Element, Name, "")
        TextNode = xmlQDBRequest.CreateNode(XmlNodeType.Text, "", "")
        TextNode.InnerText = Value
        Attribute = xmlQDBRequest.CreateAttribute(AttributeName)
        Attribute.Value = AttributeValue
        ElementNode.Attributes.Append(Attribute)

        ElementNode.AppendChild(TextNode)
        Root.AppendChild(ElementNode)
        Root = Nothing
        ElementNode = Nothing
        TextNode = Nothing
    End Sub


    Public Sub addCDATAParameter(ByRef xmlQDBRequest As XmlDocument, ByVal Name As String, ByVal Value As String)
        Dim Root As XmlElement
        Dim ElementNode As XmlNode
        Dim CDATANode As XmlNode

        Root = xmlQDBRequest.DocumentElement
        ElementNode = xmlQDBRequest.CreateNode(XmlNodeType.Element, Name, "")
        CDATANode = xmlQDBRequest.CreateNode(XmlNodeType.CDATA, "", "")
        CDATANode.InnerText = Value
        ElementNode.AppendChild(CDATANode)
        Root.AppendChild(ElementNode)
        Root = Nothing
        ElementNode = Nothing
        CDATANode = Nothing
    End Sub

    Public Sub addFieldParameter(ByRef xmlQDBRequest As XmlDocument, ByVal attrName As String, ByVal Name As String, ByVal Value As Object)
        Dim Root As XmlElement
        Dim ElementNode As XmlNode
        Dim TextNode As XmlNode
        Dim attrField As XmlAttribute
        Dim attrFileName As XmlAttribute


        Root = xmlQDBRequest.DocumentElement
        ElementNode = xmlQDBRequest.CreateNode(XmlNodeType.Element, "field", "")
        attrField = xmlQDBRequest.CreateAttribute(attrName)
        attrField.Value = Name
        Call ElementNode.Attributes.SetNamedItem(attrField)


        If TypeName(Value) = "FileStream" Then
            attrFileName = xmlQDBRequest.CreateAttribute("filename")
            attrFileName.Value = DirectCast(Value, FileStream).Name
            Call ElementNode.Attributes.SetNamedItem(attrFileName)
        End If

        TextNode = xmlQDBRequest.CreateNode(XmlNodeType.Text, "", "")
        If TypeName(Value) = "FileStream" Then
            TextNode.InnerText = fileEncode64(DirectCast(Value, FileStream))
        Else
            TextNode.InnerText = CStr(Value)
        End If
        ElementNode.AppendChild(TextNode)

        Root.AppendChild(ElementNode)
        Root = Nothing
        ElementNode = Nothing
        attrField = Nothing
        TextNode = Nothing
    End Sub
    Function int64ToDate(ByVal int64 As String) As Date
        int64ToDate = Date.FromOADate(DAYS_BETWEEN_JAVASCRIPT_AND_MICROSOFT_DATE_REFERENCES + int64toDateCommon(int64))
    End Function
    Private Function int64toDateCommon(ByVal int64 As String) As Double
        If int64 = "" Then
            Return 0
            Exit Function
        End If
        Dim dblTemp As Double
        dblTemp = makeDouble(int64)
        If dblTemp <= -59011459200001.0# Then
            dblTemp = -59011459200000.0#
        ElseIf dblTemp > 255611376000000.0# Then
            dblTemp = 255611376000000.0#
        Else
            dblTemp = (dblTemp / MILLISECONDS_IN_A_DAY)
        End If
        Return dblTemp
    End Function

    Function int64ToDuration(ByVal int64 As String) As Date
        int64ToDuration = Date.FromOADate(int64toDateCommon(int64))
    End Function

    Function makeAlphaNumLowerCase(ByVal strString As String) As String
        Dim i As Integer
        Dim chrString As String

        makeAlphaNumLowerCase = ""
        For i = 1 To Len(strString)
            chrString = Mid(strString, i, 1)
            If System.Char.IsLetterOrDigit(chrString, 0) Then
                makeAlphaNumLowerCase = makeAlphaNumLowerCase & chrString
            Else
                makeAlphaNumLowerCase = makeAlphaNumLowerCase & "_"
            End If
        Next i
        makeAlphaNumLowerCase = LCase(makeAlphaNumLowerCase)
    End Function
    Public Sub setGMTOffset(ByVal offsetHours As Single)
        GMTOffset = offsetHours
    End Sub
    Public Sub setServer(ByVal strHost As String, ByVal HTTPS As Boolean)
        If strHost <> "" Then
            QDBHost = strHost
            useHTTPS = HTTPS
        Else
            QDBHost = "www.quickbase.com"
            useHTTPS = True
        End If
    End Sub

    Public Function makeDouble(ByVal strString As String) As Double
        Dim i As Integer
        Dim chrString As String
        Dim strChar As String
        Dim resultString As String

        On Error Resume Next
        makeDouble = CDbl(strString)
        If Err.Number = 0 Then
            Exit Function
        End If
        On Error GoTo 0
        resultString = ""
        For i = 1 To Len(strString)
            strChar = Mid(strString, i, 1)
            If (((Not System.Char.IsLetter(strChar, 0)) And System.Char.IsLetterOrDigit(strChar, 0)) Or strChar = "." Or strChar = "-") Then
                resultString = resultString & strChar
            End If
        Next i
        On Error Resume Next
        makeDouble = CDbl(resultString)
        Exit Function
    End Function


    Function fileEncode64(ByVal fileToUpload As FileStream) As String
        Dim triplicate As Integer
        Dim i As Integer
        Dim outputText As String
        Dim fileLength As Integer
        Dim fileTriads As Integer
        Dim firstByte(0) As Byte
        Dim secondByte(0) As Byte
        Dim thirdByte(0) As Byte
        Dim fileRemainder As Integer

        fileLength = CInt(fileToUpload.Length)
        fileRemainder = CInt(fileLength Mod 3)
        fileTriads = fileLength \ 3
        If fileRemainder > 0 Then
            outputText = Space((fileTriads + 1) * 4)
        Else
            outputText = Space(fileTriads * 4)
        End If


        For i = 0 To fileTriads - 1             ' loop through octets
            'build 24 bit triplicate
            fileToUpload.Read(firstByte, 0, 1)
            fileToUpload.Read(secondByte, 0, 1)
            fileToUpload.Read(thirdByte, 0, 1)

            triplicate = (CInt(firstByte(0)) * 65536) + (CInt(secondByte(0)) * CInt(256)) + CInt(thirdByte(0))
            'extract four 6 bit quartets from triplicate
            Mid(outputText, (i * 4) + 1) = Mid(Map64, (triplicate \ 262144) + 1, 1) & Mid(Map64, ((triplicate And 258048) \ 4096) + 1, 1) & Mid(Map64, ((triplicate And 4032) \ 64) + 1, 1) & Mid(Map64, (triplicate And 63) + 1, 1)
        Next                                                    ' next octet
        Select Case fileRemainder
            Case 1
                fileToUpload.Read(firstByte, 0, 1)
                triplicate = (firstByte(0) * 65536)
                Mid(outputText, (i * 4) + 1) = Mid(Map64, (triplicate \ 262144) + 1, 1) & Mid(Map64, ((triplicate And 258048) \ 4096) + 1, 1) & "="
            Case 2
                fileToUpload.Read(firstByte, 0, 1)
                fileToUpload.Read(secondByte, 0, 1)
                triplicate = (firstByte(0) * 65536) + (secondByte(0) * 256)
                Mid(outputText, (i * 4) + 1) = Mid(Map64, (triplicate \ 262144) + 1, 1) & Mid(Map64, ((triplicate And 258048) \ 4096) + 1, 1) & Mid(Map64, ((triplicate And 4032) \ 64) + 1, 1) & "="
        End Select
        fileEncode64 = outputText
    End Function

    Function makeValidFilename(ByVal strString As String) As String
        Dim i As Integer
        Dim byteChar As String
        makeValidFilename = ""
        For i = 1 To Len(strString)
            byteChar = Mid(strString, i, 1)
            If byteChar = "\" Or byteChar = "/" Or
               byteChar = ":" Or byteChar = "*" Or
               Asc(byteChar) = 63 Or byteChar = """" Or
               byteChar = "<" Or byteChar = ">" Or
               byteChar = "|" Or byteChar = "'" _
            Then
                makeValidFilename = makeValidFilename & "_"
            Else
                makeValidFilename = makeValidFilename + byteChar
            End If
        Next i
    End Function

    Public Sub New(ByVal uid As String, ByVal pwd As String)
        UserName = uid
        Password = pwd
        GMTOffset = -7
    End Sub
End Class