
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
    Private qdb As QuickBaseClient
    Private Const AppName = "QuNectRestore"
    Dim fieldNodes As XmlNodeList
    Dim schemaXML As XmlDocument
    Dim clist As String()
    Dim rids As HashSet(Of String) = Nothing
    Dim keyfid As String
    Dim uniqueFieldValues As Hashtable
    Dim frmErr As New frmErr
    Private Function restoreTable(checkForErrorsOnly As Boolean) As Boolean
        restoreTable = True

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
            Dim destComboBoxCell As DataGridViewComboBoxCell = DirectCast(dgMapping.Rows(i).Cells(3), System.Windows.Forms.DataGridViewComboBoxCell)
            If destComboBoxCell.Value Is Nothing Then Continue For
            Dim destDDIndex = destComboBoxCell.Items.IndexOf(destComboBoxCell.Value)

            If destDDIndex > 0 Then
                'this is a field that needs importing
                Dim fieldNode As XmlNode = fieldNodes(destDDIndex - 1)
                fields.Add(fieldNode)
                Dim fid As String = fieldNode.SelectSingleNode("@id").InnerText
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
                fieldLabels.Add(fieldNode.SelectSingleNode("label").InnerText)
            End If
        Next
        'now we need to look for unique fields
        If Not schemaXML.SelectSingleNode("/*/table/original/key_fid") Is Nothing Then
            keyfid = schemaXML.SelectSingleNode("/*/table/original/key_fid").InnerText()
        Else
            keyfid = "3"
        End If
        'also need to pull all the fields marked required
        If checkForErrorsOnly Then
            frmErr.rdbCancel.Checked = True
            uniqueFieldValues = New Hashtable
            For i As Integer = 0 To fieldNodes.Count - 1
                Dim fid As String = fieldNodes(i).SelectSingleNode("@id").InnerText()
                If Not fieldNodes(i).SelectSingleNode("unique") Is Nothing AndAlso fieldNodes(i).SelectSingleNode("unique").InnerText() = "1" Then
                    If Not fidsForImport.Contains(fid) Then
                        requiredFieldsErrors &= "You must import values into the unique field: " & fieldNodes(i).SelectSingleNode("label").InnerText()
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
                If Not fieldNodes(i).SelectSingleNode("required") Is Nothing AndAlso fieldNodes(i).SelectSingleNode("required").InnerText() = "1" Then
                    If Not fidsForImport.Contains(fid) Then
                        requiredFieldsErrors &= "You must import values into the required field: " & fieldNodes(i).SelectSingleNode("label").InnerText()
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
        End If
        Dim connectionString As String = "Driver={QuNect ODBC for QuickBase};uid=" & txtUsername.Text & ";pwd=" & txtPassword.Text & ";QUICKBASESERVER=" & txtServer.Text & ";USEFIDS=1;APPTOKEN=" & txtAppToken.Text
        Using connection As OdbcConnection = New OdbcConnection(connectionString)
            connection.Open()
            strSQL &= "fid" & String.Join(",fid", fids.ToArray)
            strSQL &= ") VALUES ("
            strSQL &= String.Join(",", vals.ToArray)
            strSQL &= ")"
            Using command As OdbcCommand = New OdbcCommand(strSQL, connection)
                For i As Integer = 0 To fids.Count - 1
                    Dim fieldNode As XmlNode = fields(i)
                    Dim basetype As String = fieldNode.SelectSingleNode("@base_type").InnerText
                    Dim qdbType As OdbcType
                    Select Case basetype
                        Case "text"
                            qdbType = OdbcType.VarChar
                        Case "float"
                            If Not fieldNode.SelectSingleNode("decimal_places") Is Nothing AndAlso fieldNode.SelectSingleNode("decimal_places").InnerText <> "" Then
                                qdbType = OdbcType.Numeric
                            Else
                                qdbType = OdbcType.Double
                            End If

                        Case "bool"
                            qdbType = OdbcType.Bit
                        Case "int32"
                            qdbType = OdbcType.Int
                        Case "int64"
                            Select Case fieldNode.SelectSingleNode("@field_type").InnerText
                                Case "timestamp"
                                    qdbType = OdbcType.DateTime
                                Case "timeofday"
                                    qdbType = OdbcType.Time
                                Case "duration"
                                    qdbType = OdbcType.Double
                                Case Else
                                    qdbType = OdbcType.Date
                            End Select
                        Case Else
                            qdbType = OdbcType.VarChar
                    End Select
                    command.Parameters.Add("@fid" & fids(i), qdbType) 'we need to set the type here according to the field schema
                Next
                Dim csvReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(lblFile.Text)
                csvReader.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited
                csvReader.Delimiters = New String() {","}
                Dim currentRow As String()
                'Loop through all of the fields in the file. 
                Dim transaction As OdbcTransaction = Nothing
                transaction = connection.BeginTransaction()
                command.Transaction = transaction
                command.CommandType = CommandType.Text
                command.CommandTimeout = 0
                If Not csvReader.EndOfData And chkBxHeaders.Checked Then currentRow = csvReader.ReadFields()
                Dim lineCounter As Integer = 0
                Dim fileLineCounter As Integer = 0
                Try
                    pb.Visible = True
                    pb.Maximum = 1000
                    While Not csvReader.EndOfData
                        Try
                            currentRow = csvReader.ReadFields()
                            'here we need to walk down the rows of the datagrid to see if we need to exclude this row
                            Dim skipRow As Boolean = False
                            For i As Integer = 0 To dgMapping.Rows.Count - 1
                                Dim cellValue As String = currentRow(i)
                                Dim op As String = DirectCast(dgMapping.Rows(i).Cells(1), System.Windows.Forms.DataGridViewComboBoxCell).Value
                                Select Case op
                                    Case "has any value"
                                        Continue For
                                    Case "equals"
                                        If cellValue <> DirectCast(dgMapping.Rows(i).Cells(2), System.Windows.Forms.DataGridViewTextBoxCell).Value Then
                                            skipRow = True
                                            Exit For
                                        End If

                                    Case "does not equal"
                                        If cellValue = DirectCast(dgMapping.Rows(i).Cells(2), System.Windows.Forms.DataGridViewTextBoxCell).Value Then
                                            skipRow = True
                                            Exit For
                                        End If
                                    Case "contains"
                                        If Not cellValue.Contains(DirectCast(dgMapping.Rows(i).Cells(2), System.Windows.Forms.DataGridViewTextBoxCell).Value) Then
                                            skipRow = True
                                            Exit For
                                        End If
                                    Case "starts with"
                                        If Not cellValue.StartsWith(DirectCast(dgMapping.Rows(i).Cells(2), System.Windows.Forms.DataGridViewTextBoxCell).Value) Then
                                            skipRow = True
                                            Exit For
                                        End If
                                    Case "does not contain"
                                        If cellValue.Contains(DirectCast(dgMapping.Rows(i).Cells(2), System.Windows.Forms.DataGridViewTextBoxCell).Value) Then
                                            skipRow = True
                                            Exit For
                                        End If
                                    Case "does not start with"
                                        If cellValue.StartsWith(DirectCast(dgMapping.Rows(i).Cells(2), System.Windows.Forms.DataGridViewTextBoxCell).Value) Then
                                            skipRow = True
                                            Exit For
                                        End If
                                End Select
                            Next
                            If skipRow Then
                                Continue While
                            End If
                            'This is a row we need to import
                            For i As Integer = 0 To importThese.Count - 1
                                Dim val As String = currentRow(importThese(i))
                                If checkForErrorsOnly Then
                                    If fids(i) = "3" AndAlso Not rids.Contains(val) Then
                                        'here we're going to have a problem because this record no longer exists in QuickBase
                                        missingRIDs &= vbCrLf & "line " & fileLineCounter + 1 & " has Record ID# " & val & " which no longer exists"
                                    End If
                                    If requiredFidsForImport.Contains(fids(i)) Then
                                        If val = "" Then
                                            requiredFieldsErrors &= vbCrLf & "line " & fileLineCounter + 1 & " has a blank value mapped to required field " & fieldLabels(i)
                                            If requiredFieldsErrors.Length > 1000 Then
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
                                            uniqueFieldErrors &= vbCrLf & "line " & fileLineCounter + 1 & " has a duplicate value for unique field " & fieldLabels(i)
                                        Else
                                            uniqueFieldValues(fids(i)).add(val)
                                        End If
                                    End If
                                ElseIf fids(i) = "3" AndAlso Not rids.Contains(val) Then
                                    If frmErr.rdbSkipRecords.Checked Then
                                        fileLineCounter += 1
                                        Continue While
                                    Else
                                        val = ""
                                    End If
                                End If
                                Dim qdbType As OdbcType = command.Parameters("@fid" & fids(i)).ODBCType
                                Try
                                    Select Case qdbType
                                        Case OdbcType.Int
                                            command.Parameters("@fid" & fids(i)).Value = Convert.ToInt32(val)
                                        Case OdbcType.Double, OdbcType.Numeric
                                            command.Parameters("@fid" & fids(i)).Value = Convert.ToDouble(val)
                                        Case OdbcType.Date
                                            command.Parameters("@fid" & fids(i)).Value = Date.Parse(val)
                                        Case OdbcType.DateTime
                                            command.Parameters("@fid" & fids(i)).Value = DateTime.Parse(val)
                                        Case OdbcType.Time
                                            command.Parameters("@fid" & fids(i)).Value = TimeSpan.Parse(val)
                                        Case OdbcType.Bit
                                            If Regex.IsMatch(val, "y|1|c", RegexOptions.IgnoreCase) Then
                                                command.Parameters("@fid" & fids(i)).Value = True
                                            Else
                                                command.Parameters("@fid" & fids(i)).Value = False
                                            End If
                                        Case Else
                                            command.Parameters("@fid" & fids(i)).Value = val
                                    End Select
                                Catch ex As Exception
                                    If checkForErrorsOnly And val <> "" Then
                                        conversionErrors &= vbCrLf & "line " & fileLineCounter + 1 & " '" & val & "' to " & qdbType.ToString() & " for field " & fieldLabels(i)
                                        If conversionErrors.Length > 1000 Then
                                            Exit While
                                        End If
                                    End If
                                    If frmErr.rdbSkipRecords.Checked Then
                                        fileLineCounter += 1
                                        Continue While
                                    End If
                                    If Not checkForErrorsOnly Then
                                        Select Case qdbType
                                            Case OdbcType.Int, OdbcType.Double, OdbcType.Numeric
                                                Dim nullDouble As Double
                                                command.Parameters("@fid" & fids(i)).Value = nullDouble
                                            Case OdbcType.Date
                                                Dim nulldate As Date
                                                command.Parameters("@fid" & fids(i)).Value = nulldate
                                            Case OdbcType.DateTime
                                                Dim nullDateTime As DateTime
                                                command.Parameters("@fid" & fids(i)).Value = nullDateTime
                                            Case OdbcType.Time
                                                Dim nullTimeSpan As TimeSpan
                                                command.Parameters("@fid" & fids(i)).Value = nullTimeSpan
                                            Case OdbcType.Bit
                                                command.Parameters("@fid" & fids(i)).Value = False
                                            Case Else
                                                command.Parameters("@fid" & fids(i)).Value = ""
                                        End Select
                                    End If
                                End Try

                            Next
                            If Not checkForErrorsOnly Then command.ExecuteNonQuery()
                            lineCounter += 1
                        Catch ex As Exception
                            If checkForErrorsOnly Then
                                improperlyFormattedLines &= vbCrLf & "line " & fileLineCounter + 1 & " is not a properly formatted CSV line."
                                If improperlyFormattedLines.Length > 1000 Then
                                    Exit While
                                End If
                            End If
                        End Try
                        fileLineCounter += 1
                        pb.Value = fileLineCounter Mod 1000
                    End While
                    If Not checkForErrorsOnly Then
                        transaction.Commit()
                        MsgBox("Imported " & lineCounter & " records!", MsgBoxStyle.OkOnly, AppName)
                    Else
                        If conversionErrors <> "" Or improperlyFormattedLines <> "" Or uniqueFieldErrors <> "" Or requiredFieldsErrors <> "" Or missingRIDs <> "" Then
                            frmErr.lblMalformed.Text = improperlyFormattedLines
                            frmErr.lblUnique.Text = uniqueFieldErrors
                            frmErr.lblRequired.Text = requiredFieldsErrors
                            frmErr.lblConversions.Text = conversionErrors
                            frmErr.lblMissing.Text = missingRIDs
                            If improperlyFormattedLines <> "" Then
                                frmErr.TabControlErrors.SelectedIndex = 3
                                frmErr.TabMalformed.Text = "*Malformed lines"
                            Else
                                frmErr.TabMalformed.Text = "Malformed lines"
                            End If
                            If uniqueFieldErrors <> "" Then
                                frmErr.TabControlErrors.SelectedIndex = 2
                                frmErr.TabUnique.Text = "*Unique"
                            Else
                                frmErr.TabUnique.Text = "Unique"
                            End If
                            If requiredFieldsErrors <> "" Then
                                frmErr.TabControlErrors.SelectedIndex = 1
                                frmErr.TabRequired.Text = "*Required"
                            Else
                                frmErr.TabRequired.Text = "Required"
                            End If
                            If missingRIDs <> "" Then
                                frmErr.TabControlErrors.SelectedIndex = 0
                                frmErr.TabMissing.Text = "*Missing Record ID#s"
                            Else
                                frmErr.TabMissing.Text = "Missing Record ID#s"
                            End If
                            If conversionErrors <> "" Then
                                frmErr.TabControlErrors.SelectedIndex = 0
                                frmErr.TabConversions.Text = "*Conversion"
                            Else
                                frmErr.TabConversions.Text = "Conversion"
                            End If
                            Dim dr As DialogResult = frmErr.ShowDialog()
                            If frmErr.rdbCancel.Checked Then
                                restoreTable = False
                            Else
                                restoreTable = True
                            End If
                        End If
                    End If
                    pb.Visible = False
                Catch ex As Exception
                    pb.Visible = False
                    MsgBox("Could not import because " & ex.Message)
                    restoreTable = False
                End Try
            End Using
        End Using
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
        MsgBox("Operation timed out. Please try again.")
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
    Sub schemaCallback(ByVal result As System.IAsyncResult)
        Try
            Dim requestState As QuickBaseClient.WebRequestState = CType(result.AsyncState, QuickBaseClient.WebRequestState)
            schemaXML = requestState.xmlDoc
            fieldNodes = schemaXML.SelectNodes("/*/table/fields/field[(not(@mode) and not(@role)) or @id=3]")


            Dim fidReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(Regex.Replace(lblFile.Text, "csv$", "fids"))


            fidReader.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited
            fidReader.Delimiters = New String() {"."}
            clist = fidReader.ReadFields()

            'here we need to open the csv and get the field names
            Dim csvReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(lblFile.Text)
            csvReader.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited
            csvReader.Delimiters = New String() {","}
            Dim currentRow As String()
            'Loop through all of the fields in the schema.
            DirectCast(dgMapping.Columns(3), System.Windows.Forms.DataGridViewComboBoxColumn).Items.Clear()
            DirectCast(dgMapping.Columns(3), System.Windows.Forms.DataGridViewComboBoxColumn).Items.Add("Do not import")
            For i As Integer = 0 To fieldNodes.Count - 1
                DirectCast(dgMapping.Columns(3), System.Windows.Forms.DataGridViewComboBoxColumn).Items.Add(fieldNodes(i).SelectSingleNode("label").InnerText)
            Next
            If Not csvReader.EndOfData Then
                Try
                    currentRow = csvReader.ReadFields()
                    dgMapping.Rows.Clear()
                    For i As Integer = 0 To currentRow.Length - 1
                        If currentRow(i) = "" Then Continue For
                        Dim j As Integer = dgMapping.Rows.Add(New String() {currentRow(i)})
                        DirectCast(dgMapping.Rows(j).Cells(1), System.Windows.Forms.DataGridViewComboBoxCell).Value = "has any value"
                        'DirectCast(dgMapping.Rows(j).Cells(3), System.Windows.Forms.DataGridViewComboBoxCell).Value = "Do not import"
                        'now we have to see if we can match this to a field in the table
                        If clist.Length > i + 1 Then
                            For k As Integer = 0 To fieldNodes.Count - 1
                                If fieldNodes(k).SelectSingleNode("@id").InnerText = clist(i) Then
                                    DirectCast(dgMapping.Rows(j).Cells(3), System.Windows.Forms.DataGridViewComboBoxCell).Value = fieldNodes(k).SelectSingleNode("label").InnerText
                                    Exit For
                                End If
                            Next
                        Else
                            For k As Integer = 0 To fieldNodes.Count - 1
                                If fieldNodes(k).SelectSingleNode("label").InnerText = currentRow(i) Then
                                    DirectCast(dgMapping.Rows(j).Cells(3), System.Windows.Forms.DataGridViewComboBoxCell).Value = currentRow(i)
                                    Exit For
                                End If
                            Next
                        End If
                    Next

                Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                    MsgBox("First line of file is malformed, cannot display field names.")
                End Try
            End If
            btnImport.Visible = True
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
        Dim ops As String() = New String() {"has any value", "equals", "does not equal", "contains", "does not contain", "starts with", "does not start with"}
        For i As Integer = 0 To ops.Count - 1
            DirectCast(dgMapping.Columns(1), System.Windows.Forms.DataGridViewComboBoxColumn).Items.Add(ops(i))
        Next
    End Sub

    Private Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click
        If restoreTable(True) Then
            restoreTable(False)
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
        qdb.setServer(txtServer.Text, True)
        qdb.setAppToken(txtAppToken.Text)
        qdb.Authenticate(txtUsername.Text, txtPassword.Text)
        qdb.GetSchema(lblTable.Text, New AsyncCallback(AddressOf schemaCallback), New AsyncCallback(AddressOf timeoutCallback))
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

    Private Sub cmbAttachments_SelectedIndexChanged(sender As Object, e As EventArgs)

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
            fsWrite = New FileStream(Path:=fileName, _
                mode:=FileMode.CreateNew, access:=FileAccess.Write)

            'create binary writer instance
            Dim bWrite As BinaryWriter
            bWrite = New BinaryWriter(output:=fsWrite)
            'write bytes out 
            bWrite.Write(byteResponseArray, 0, byteResponseArray.Length)


            'close the writer 
            bWrite.Close()

            fsWrite.Close()


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
            Exit Function
        End If
        Dim dblTemp As Double
        dblTemp = makeDouble(int64)
        If dblTemp <= -59011459200001.0# Then
            dblTemp = -59011459200000.0#
        ElseIf dblTemp > 255611376000000.0# Then
            dblTemp = 255611376000000.0#
        Else
            int64toDateCommon = (dblTemp / MILLISECONDS_IN_A_DAY)
        End If
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
            If byteChar = "\" Or byteChar = "/" Or _
               byteChar = ":" Or byteChar = "*" Or _
               Asc(byteChar) = 63 Or byteChar = """" Or _
               byteChar = "<" Or byteChar = ">" Or _
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