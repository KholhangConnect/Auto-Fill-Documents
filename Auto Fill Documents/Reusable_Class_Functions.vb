Public Class Reusable_Class_Functions
    ' Constants
    Private Const DEFAULT_FOLDER As String = "AUTOFILL_DOCS"

    ' COM Object Cleanup
    Public Shared Sub ReleaseComObject(obj As Object)
        Try
            If obj IsNot Nothing Then
                While System.Runtime.InteropServices.Marshal.ReleaseComObject(obj) > 0
                    ' Continue releasing until reference count is 0
                End While
                obj = Nothing
            End If
        Catch
            obj = Nothing
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub

    ' File System Operations
    Public Shared Function CanCreateDirectory(path As String) As Boolean
        Try
            ' Check if we have permission to create directory
            Dim securityTest = IO.Directory.CreateDirectory(path)
            securityTest.Delete()
            Return True
        Catch
            Return False
        End Try
    End Function

    Public Shared Function GetDefaultOutputPath() As String
        Try
            ' Get all available drives
            Dim drives = IO.DriveInfo.GetDrives().Where(Function(d) d.IsReady).ToList()

            ' Try to find D: drive first
            Dim targetDrive = drives.FirstOrDefault(Function(d) d.Name.StartsWith("D"))

            ' If D: drive not found, use C: drive
            If targetDrive Is Nothing Then
                targetDrive = drives.FirstOrDefault(Function(d) d.Name.StartsWith("C"))
            End If

            If targetDrive IsNot Nothing Then
                ' Create base path (Drive:\AUTOFILL_DOCS)
                Dim basePath = IO.Path.Combine(targetDrive.RootDirectory.FullName, DEFAULT_FOLDER)

                ' If that's not available, use Documents folder
                If Not IO.Directory.Exists(basePath) AndAlso Not CanCreateDirectory(basePath) Then
                    basePath = IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), DEFAULT_FOLDER)
                End If

                ' Create directory if it doesn't exist
                If Not IO.Directory.Exists(basePath) Then
                    IO.Directory.CreateDirectory(basePath)
                End If

                Return basePath
            Else
                ' Fallback to Documents folder if no suitable drive found
                Dim docPath = IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), DEFAULT_FOLDER)
                If Not IO.Directory.Exists(docPath) Then
                    IO.Directory.CreateDirectory(docPath)
                End If
                Return docPath
            End If

        Catch ex As Exception
            ' If anything fails, use Documents folder as fallback
            Try
                Dim fallbackPath = IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), DEFAULT_FOLDER)
                If Not IO.Directory.Exists(fallbackPath) Then
                    IO.Directory.CreateDirectory(fallbackPath)
                End If
                Return fallbackPath
            Catch innerEx As Exception
                Throw New Exception("Error setting default path: " & innerEx.Message)
            End Try
        End Try
    End Function

    ' Excel Operations
    Public Shared Function GetExcelHeaders(excelPath As String) As List(Of String)
        If Not IO.File.Exists(excelPath) Then
            Throw New Exception($"Excel file not found: {excelPath}")
        End If

        Dim excel As Object = Nothing
        Dim workbook As Object = Nothing
        Dim worksheet As Object = Nothing
        Dim usedRange As Object = Nothing
        Dim headers As New List(Of String)

        Try
            ' Create Excel Application instance
            excel = CreateObject("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            ' Open workbook
            workbook = excel.Workbooks.Open(excelPath)
            worksheet = workbook.Worksheets(1)
            usedRange = worksheet.UsedRange

            ' Get column headers
            Dim columnCount = usedRange.Columns.Count
            For i = 1 To columnCount
                Dim header = CStr(worksheet.Cells(1, i).Value).Trim()
                If Not String.IsNullOrEmpty(header) Then
                    headers.Add(header)
                End If
            Next

            Return headers

        Catch ex As Exception
            Throw New Exception($"Error reading Excel file: {ex.Message}")
        Finally
            ' Cleanup in reverse order
            If usedRange IsNot Nothing Then ReleaseComObject(usedRange)
            If worksheet IsNot Nothing Then ReleaseComObject(worksheet)
            If workbook IsNot Nothing Then
                workbook.Close(False)
                ReleaseComObject(workbook)
            End If
            If excel IsNot Nothing Then
                excel.Quit()
                ReleaseComObject(excel)
            End If
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Function

    Public Shared Function GetColumnIndex(worksheet As Object, headerName As String) As Integer
        Dim range = worksheet.UsedRange
        For i = 1 To range.Columns.Count
            If CStr(worksheet.Cells(1, i).Value) = headerName Then
                Return i
            End If
        Next
        Return -1
    End Function

    Public Shared Function FindMatchingRow(worksheet As Object, joinColumnIndex As Integer, searchValue As String) As Integer
        Dim range = worksheet.UsedRange
        For row = 2 To range.Rows.Count ' Start from 2 to skip header
            If CStr(worksheet.Cells(row, joinColumnIndex).Value) = searchValue Then
                Return row
            End If
        Next
        Return -1
    End Function

    ' Word Operations
    Public Shared Function ExtractFieldNamesFromWord(filePath As String) As List(Of String)
        If Not IO.File.Exists(filePath) Then
            Throw New Exception($"Word file not found: {filePath}")
        End If

        Dim word As Object = Nothing
        Dim doc As Object = Nothing
        Dim range As Object = Nothing
        Dim fieldNames As New HashSet(Of String)  ' Using HashSet to automatically handle duplicates

        Try
            ' Create Word Application instance
            word = CreateObject("Word.Application")
            word.Visible = False

            ' Open document
            doc = word.Documents.Open(filePath)
            range = doc.Range()

            ' Get the document text
            Dim content As String = range.Text

            ' Find all matches of <<field_name>> pattern
            Dim pattern As String = "<<([^<>]+)>>"
            Dim matches = System.Text.RegularExpressions.Regex.Matches(content, pattern)

            ' Extract field names from matches
            For Each match As System.Text.RegularExpressions.Match In matches
                If match.Groups.Count > 1 Then
                    Dim fieldName = match.Groups(1).Value.Trim()
                    If Not String.IsNullOrEmpty(fieldName) Then
                        fieldNames.Add(fieldName)
                    End If
                End If
            Next

            Return fieldNames.ToList()

        Catch ex As Exception
            Throw New Exception($"Error reading Word file: {ex.Message}")
        Finally
            ' Cleanup in reverse order
            If range IsNot Nothing Then ReleaseComObject(range)
            If doc IsNot Nothing Then
                doc.Close(False)
                ReleaseComObject(doc)
            End If
            If word IsNot Nothing Then
                word.Quit()
                ReleaseComObject(word)
            End If
        End Try
    End Function

    ' HTML Generation
    Public Shared Function GenerateHtmlHeader(title As String) As Text.StringBuilder
        Dim sb As New Text.StringBuilder()
        sb.AppendLine("<!DOCTYPE html>")
        sb.AppendLine("<html lang='en'>")
        sb.AppendLine("<head>")
        sb.AppendLine("    <meta charset='UTF-8'>")
        sb.AppendLine("    <meta name='viewport' content='width=device-width, initial-scale=1.0'>")
        sb.AppendLine($"    <title>{title}</title>")
        sb.AppendLine("    <style>")
        sb.AppendLine("        body {")
        sb.AppendLine("            font-family: Arial, sans-serif;")
        sb.AppendLine("            line-height: 1.6;")
        sb.AppendLine("            max-width: 1200px;")
        sb.AppendLine("            margin: 0 auto;")
        sb.AppendLine("            padding: 20px;")
        sb.AppendLine("            background-color: #f5f5f5;")
        sb.AppendLine("        }")
        sb.AppendLine("        .container {")
        sb.AppendLine("            background-color: white;")
        sb.AppendLine("            padding: 30px;")
        sb.AppendLine("            border-radius: 8px;")
        sb.AppendLine("            box-shadow: 0 2px 4px rgba(0,0,0,0.1);")
        sb.AppendLine("        }")
        sb.AppendLine("        h1 {")
        sb.AppendLine("            color: #2c3e50;")
        sb.AppendLine("            text-align: center;")
        sb.AppendLine("            border-bottom: 2px solid #3498db;")
        sb.AppendLine("            padding-bottom: 10px;")
        sb.AppendLine("        }")
        sb.AppendLine("        h2 {")
        sb.AppendLine("            color: #3498db;")
        sb.AppendLine("            margin-top: 20px;")
        sb.AppendLine("        }")
        sb.AppendLine("        .footer {")
        sb.AppendLine("            text-align: center;")
        sb.AppendLine("            margin-top: 30px;")
        sb.AppendLine("            color: #7f8c8d;")
        sb.AppendLine("            font-size: 0.9em;")
        sb.AppendLine("        }")
        sb.AppendLine("    </style>")
        sb.AppendLine("</head>")
        Return sb
    End Function

    Public Shared Function GenerateHtmlFooter(Optional includeVersion As Boolean = False) As String
        Dim footer As New Text.StringBuilder()
        footer.AppendLine("        <div class='footer'>")
        footer.AppendLine($"            <p>Generated on: {DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}</p>")
        If includeVersion Then
            footer.AppendLine("            <p>Auto Fill Documents Application - Version 1.0</p>")
        End If
        footer.AppendLine("        </div>")
        footer.AppendLine("    </div>")
        footer.AppendLine("</body>")
        footer.AppendLine("</html>")
        Return footer.ToString()
    End Function

    ' File Opening
    Public Shared Sub OpenFileInBrowser(filePath As String)
        Try
            Dim psi As New ProcessStartInfo()
            psi.FileName = filePath
            psi.UseShellExecute = True
            psi.Verb = "open"
            Process.Start(psi)
        Catch ex As Exception
            Throw New Exception($"Error opening file: {ex.Message}")
        End Try
    End Sub

    ' Progress Updates
    Public Shared Sub UpdateProgressBar(progressBar As ProgressBar, label As Label, percentage As Integer, status As String)
        If progressBar IsNot Nothing Then
            progressBar.Value = percentage
        End If
        If label IsNot Nothing Then
            label.Text = status
        End If
        Application.DoEvents()
    End Sub

    ' Form Control Management
    Public Shared Sub InitializeComboBoxes(ParamArray comboBoxes() As ComboBox)
        For Each cb In comboBoxes
            cb.DropDownStyle = ComboBoxStyle.DropDownList
            cb.BackColor = Color.White
        Next
    End Sub

    Public Shared Function AreAllCheckedInGroup(checkBoxes() As CheckBox) As Boolean
        Return checkBoxes.All(Function(cb) cb.Checked)
    End Function

    Public Shared Sub HandleCheckBoxChange(checkBox As CheckBox, button As Button, textBox As TextBox,
                                         groupCheckBoxes() As CheckBox, groupNumber As Integer,
                                         Optional updateGroupBox4Action As Action = Nothing,
                                         Optional updateComboBoxAction As Action = Nothing)
        Try
            If checkBox.Checked Then
                ' Check if this would make all checkboxes checked in the group
                If AreAllCheckedInGroup(groupCheckBoxes) Then
                    checkBox.Checked = False
                    MessageBox.Show("At least one document/file must not be ignored in each group.",
                                  "Invalid Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If
                textBox.Text = ""
            End If

            button.Enabled = Not checkBox.Checked

            ' Update GroupBox4 state if this is an Excel checkbox (group 2)
            If groupNumber = 2 Then
                updateGroupBox4Action?.Invoke()
                updateComboBoxAction?.Invoke()
            End If

        Catch ex As Exception
            MessageBox.Show($"Error handling checkbox change: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' File Dialog Operations
    Public Shared Sub ShowFileDialog(filter As String, title As String, textBox As TextBox,
                                   checkBoxes() As CheckBox, textBoxes() As TextBox,
                                   Optional updateGroupBox4Action As Action = Nothing)
        Try
            Using openFileDialog As New OpenFileDialog()
                openFileDialog.Filter = filter
                openFileDialog.Title = title

                If openFileDialog.ShowDialog() = DialogResult.OK Then
                    ' Check for duplicate file selection
                    Dim selectedPath = openFileDialog.FileName
                    Dim duplicateInfo = CheckDuplicateFile(selectedPath, checkBoxes, textBoxes, textBox)

                    If duplicateInfo.Item1 Then
                        MessageBox.Show($"This file has already been selected in {duplicateInfo.Item2}. Please either:" & vbCrLf & vbCrLf &
                                      "1. Select a different file, or" & vbCrLf &
                                      "2. Check the 'Ignore' checkbox for the current selection if you don't need it.",
                                      "Duplicate File Selection",
                                      MessageBoxButtons.OK,
                                      MessageBoxIcon.Information)
                        Return
                    End If

                    textBox.Text = selectedPath
                    MessageBox.Show("The file will be validate! and may take less than 5sec or depending on the file records uploaded for processing",
                                      "Validating File uploaded",
                                      MessageBoxButtons.OK,
                                      MessageBoxIcon.Information)

                    ' Update GroupBox4 state if this is an Excel file
                    If filter.Contains("Excel") Then
                        updateGroupBox4Action?.Invoke()
                    End If
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show("Error opening file dialog: " & ex.Message, "File Dialog Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Shared Function CheckDuplicateFile(selectedPath As String,
                                             checkBoxes() As CheckBox,
                                             textBoxes() As TextBox,
                                             currentTextBox As TextBox) As Tuple(Of Boolean, String)
        For i As Integer = 0 To textBoxes.Length - 1
            If Not checkBoxes(i).Checked AndAlso
               textBoxes(i).Text = selectedPath AndAlso
               textBoxes(i) IsNot currentTextBox Then
                Return Tuple.Create(True, $"File {i + 1}")
            End If
        Next
        Return Tuple.Create(False, "")
    End Function

    ' Field Matching Dialog
    Public Shared Function ShowFieldMatchingDialog(wordFields As List(Of String),
                                                 excelHeaders As List(Of String),
                                                 documentNumber As Integer) As Dictionary(Of String, String)
        ' Create a new form for field matching
        Dim matchForm As New Form()
        matchForm.Text = $"Match Fields - Document {documentNumber}"
        matchForm.Size = New Size(400, 500)
        matchForm.StartPosition = FormStartPosition.CenterParent
        matchForm.MinimizeBox = False
        matchForm.MaximizeBox = False
        matchForm.FormBorderStyle = FormBorderStyle.FixedDialog

        ' Create layout panel
        Dim panel As New TableLayoutPanel()
        panel.Dock = DockStyle.Fill
        panel.ColumnCount = 2
        panel.RowCount = wordFields.Count + 2
        panel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))
        panel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))
        panel.Padding = New Padding(10)

        ' Add headers
        AddHeadersToPanel(panel, documentNumber)

        ' Dictionary to store ComboBoxes
        Dim fieldMatches As New Dictionary(Of String, ComboBox)

        ' Add field matching controls
        AddFieldMatchingControls(panel, wordFields, excelHeaders, fieldMatches)

        ' Add OK button
        AddOKButton(panel, wordFields.Count)

        matchForm.Controls.Add(panel)

        ' Show dialog and process results
        Return ProcessDialogResult(matchForm, fieldMatches)
    End Function

    Private Shared Sub AddHeadersToPanel(panel As TableLayoutPanel, documentNumber As Integer)
        Dim headerLabel1 As New Label()
        headerLabel1.Text = $"Word Fields (Document {documentNumber})"
        headerLabel1.Font = New Font(headerLabel1.Font, FontStyle.Bold)
        panel.Controls.Add(headerLabel1, 0, 0)

        Dim headerLabel2 As New Label()
        headerLabel2.Text = "Excel Headers"
        headerLabel2.Font = New Font(headerLabel2.Font, FontStyle.Bold)
        panel.Controls.Add(headerLabel2, 1, 0)
    End Sub

    Private Shared Sub AddFieldMatchingControls(panel As TableLayoutPanel,
                                              wordFields As List(Of String),
                                              excelHeaders As List(Of String),
                                              fieldMatches As Dictionary(Of String, ComboBox))
        For i As Integer = 0 To wordFields.Count - 1
            ' Add Word field label
            Dim label As New Label()
            label.Text = wordFields(i)
            label.Dock = DockStyle.Fill
            label.TextAlign = ContentAlignment.MiddleLeft
            panel.Controls.Add(label, 0, i + 1)

            ' Add Excel header ComboBox
            Dim combo As New ComboBox()
            combo.Dock = DockStyle.Fill
            combo.DropDownStyle = ComboBoxStyle.DropDownList
            combo.Items.AddRange(excelHeaders.ToArray())
            combo.Items.Add("(None)")
            combo.SelectedItem = "(None)"
            panel.Controls.Add(combo, 1, i + 1)

            fieldMatches.Add(wordFields(i), combo)
        Next
    End Sub

    Private Shared Sub AddOKButton(panel As TableLayoutPanel, lastRow As Integer)
        Dim btnOK As New Button()
        btnOK.Text = "OK"
        btnOK.DialogResult = DialogResult.OK
        btnOK.Dock = DockStyle.Bottom
        panel.Controls.Add(btnOK, 0, lastRow + 1)
        panel.SetColumnSpan(btnOK, 2)
    End Sub

    Private Shared Function ProcessDialogResult(matchForm As Form,
                                              fieldMatches As Dictionary(Of String, ComboBox)) As Dictionary(Of String, String)
        Dim result As New Dictionary(Of String, String)
        If matchForm.ShowDialog() = DialogResult.OK Then
            For Each kvp In fieldMatches
                Dim selectedValue As String = kvp.Value.SelectedItem.ToString()
                If selectedValue <> "(None)" Then
                    result.Add(kvp.Key, selectedValue)
                End If
            Next
        End If
        Return result
    End Function

    ' Document Processing
    Public Shared Sub ProcessWordDocument(wordPath As String, word As Object, worksheet As Object,
                                        rowIndex As Integer, headers As List(Of String),
                                        docPrefix As String, outputFolder As String,
                                        Optional fieldMappings As Dictionary(Of String, String) = Nothing)
        Dim doc As Object = Nothing
        Dim findObject As Object = Nothing

        Try
            doc = word.Documents.Open(wordPath)
            findObject = word.Selection.Find

            If fieldMappings IsNot Nothing Then
                ProcessFieldMappings(findObject, fieldMappings, worksheet, rowIndex, headers)
            Else
                ProcessDirectFieldMatching(findObject, worksheet, rowIndex, headers)
            End If

            ' Save as new document
            Dim outputFileName = $"{docPrefix}_Row{rowIndex - 1}{IO.Path.GetExtension(wordPath)}"
            Dim outputPath = IO.Path.Combine(outputFolder, outputFileName)
            doc.SaveAs(outputPath)

        Finally
            If findObject IsNot Nothing Then ReleaseComObject(findObject)
            If doc IsNot Nothing Then
                doc.Close(False)
                ReleaseComObject(doc)
            End If
        End Try
    End Sub

    Private Shared Sub ProcessFieldMappings(findObject As Object, fieldMappings As Dictionary(Of String, String),
                                          worksheet As Object, rowIndex As Integer, headers As List(Of String))
        For Each mapping In fieldMappings
            Dim columnIndex = headers.IndexOf(mapping.Value) + 1
            If columnIndex > 0 Then
                Dim cellValue = worksheet.Cells(rowIndex, columnIndex).Value
                Dim replacementText = If(cellValue IsNot Nothing, CStr(cellValue), "")
                ExecuteFindAndReplace(findObject, mapping.Key, replacementText)
            End If
        Next
    End Sub

    Private Shared Sub ProcessDirectFieldMatching(findObject As Object, worksheet As Object,
                                                rowIndex As Integer, headers As List(Of String))
        For columnIndex = 1 To headers.Count
            Dim fieldName = headers(columnIndex - 1)
            Dim cellValue = worksheet.Cells(rowIndex, columnIndex).Value
            Dim replacementText = If(cellValue IsNot Nothing, CStr(cellValue), "")
            ExecuteFindAndReplace(findObject, fieldName, replacementText)
        Next
    End Sub

    Private Shared Sub ExecuteFindAndReplace(findObject As Object, fieldName As String, replacementText As String)
        With findObject
            .Text = $"<<{fieldName}>>"
            .Replacement.Text = replacementText
            .Forward = True
            .Wrap = 0 ' wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute(Replace:=2) ' wdReplaceAll
        End With
    End Sub

    ' Data Joining Operations
    Public Shared Function JoinExcelData(activeFiles As List(Of Tuple(Of String, ComboBox, ComboBox))) As DataTable
        Dim resultTable As New DataTable("JoinedData")
        Dim excel As Object = Nothing
        Dim workbooks As New List(Of Object)()
        Dim worksheets As New List(Of Object)()

        Try
            ' Create Excel Application instance
            excel = CreateObject("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            If activeFiles.Count < 2 Then
                MessageBox.Show("At least two Excel files must be selected for joining data.",
                              "Invalid Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return Nothing
            End If

            ' Load all Excel files and their data
            For Each fileInfo In activeFiles
                Dim workbook = excel.Workbooks.Open(fileInfo.Item1)
                workbooks.Add(workbook)
                Dim worksheet = workbook.Worksheets(1)
                worksheets.Add(worksheet)

                ' Get headers from first row
                Dim usedRange = worksheet.UsedRange
                Dim headers = New List(Of String)()
                For i = 1 To usedRange.Columns.Count
                    headers.Add(CStr(worksheet.Cells(1, i).Value))
                Next

                ' If this is the first file, initialize the result table
                If resultTable.Columns.Count = 0 Then
                    For Each header In headers
                        resultTable.Columns.Add(header, GetType(String))
                    Next
                End If
            Next

            ' Perform the join operation
            resultTable = PerformJoinOperation(worksheets, activeFiles, resultTable)

            Return resultTable

        Catch ex As Exception
            MessageBox.Show($"Error joining Excel data: {ex.Message}", "Join Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        Finally
            ' Cleanup
            CleanupExcelObjects(worksheets, workbooks, excel)
        End Try
    End Function

    Private Shared Function PerformJoinOperation(worksheets As List(Of Object),
                                               activeFiles As List(Of Tuple(Of String, ComboBox, ComboBox)),
                                               resultTable As DataTable) As DataTable
        Dim baseWorksheet = worksheets(0)
        Dim baseRange = baseWorksheet.UsedRange
        Dim baseRows = baseRange.Rows.Count

        ' Get join field indices
        Dim joinFields = GetJoinFieldIndices(worksheets, activeFiles)
        If joinFields Is Nothing Then Return Nothing

        ' Process each row in base worksheet
        For baseRow = 2 To baseRows
            Dim baseKey = CStr(baseWorksheet.Cells(baseRow, joinFields(0)).Value)
            If ProcessMatchingRows(baseKey, worksheets, joinFields, baseWorksheet, baseRange, baseRow, resultTable) Then
                Continue For
            End If
        Next

        Return resultTable
    End Function

    Private Shared Function GetJoinFieldIndices(worksheets As List(Of Object),
                                              activeFiles As List(Of Tuple(Of String, ComboBox, ComboBox))) As List(Of Integer)
        Dim joinFields As New List(Of Integer)
        For i = 0 To activeFiles.Count - 1
            Dim joinFieldName = If(activeFiles(i).Item3.SelectedItem, "")
            If String.IsNullOrEmpty(joinFieldName.ToString()) Then
                MessageBox.Show($"Please select a join field for Excel file {i + 1}",
                              "Join Field Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return Nothing
            End If
            joinFields.Add(GetColumnIndex(worksheets(i), joinFieldName.ToString()))
        Next
        Return joinFields
    End Function

    Private Shared Function ProcessMatchingRows(baseKey As String, worksheets As List(Of Object),
                                              joinFields As List(Of Integer), baseWorksheet As Object,
                                              baseRange As Object, baseRow As Integer,
                                              resultTable As DataTable) As Boolean
        ' Check for matches in all worksheets
        For i = 1 To worksheets.Count - 1
            Dim worksheet = worksheets(i)
            Dim matchRow = FindMatchingRow(worksheet, joinFields(i), baseKey)
            If matchRow = -1 Then Return True
        Next

        ' Add combined row to result
        Dim newRow = resultTable.NewRow()

        ' Add data from base worksheet
        For col = 1 To baseRange.Columns.Count
            Dim header = CStr(baseWorksheet.Cells(1, col).Value)
            newRow(header) = CStr(baseWorksheet.Cells(baseRow, col).Value)
        Next

        ' Add data from other worksheets
        For i = 1 To worksheets.Count - 1
            Dim worksheet = worksheets(i)
            Dim range = worksheet.UsedRange
            Dim matchRow = FindMatchingRow(worksheet, joinFields(i), baseKey)

            For col = 1 To range.Columns.Count
                Dim header = CStr(worksheet.Cells(1, col).Value)
                If Not resultTable.Columns.Contains(header) Then
                    resultTable.Columns.Add(header, GetType(String))
                End If
                newRow(header) = CStr(worksheet.Cells(matchRow, col).Value)
            Next
        Next

        resultTable.Rows.Add(newRow)
        Return False
    End Function

    Private Shared Sub CleanupExcelObjects(worksheets As List(Of Object),
                                         workbooks As List(Of Object),
                                         excel As Object)
        For Each worksheet In worksheets
            If worksheet IsNot Nothing Then ReleaseComObject(worksheet)
        Next
        For Each workbook In workbooks
            If workbook IsNot Nothing Then
                workbook.Close(False)
                ReleaseComObject(workbook)
            End If
        Next
        If excel IsNot Nothing Then
            excel.Quit()
            ReleaseComObject(excel)
        End If
    End Sub

    ' HTML Generation
    Public Shared Sub ShowJoinedDataPreview(joinedData As DataTable)
        If joinedData Is Nothing OrElse joinedData.Rows.Count = 0 Then
            MessageBox.Show("No matching records found after joining the data.",
                          "No Data", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Try
            ' Generate HTML content
            Dim htmlContent = GenerateJoinedDataHtml(joinedData)

            ' Save and show the HTML file
            Dim tempPath As String = Environment.GetEnvironmentVariable("TEMP")
            Dim htmlFile As String = IO.Path.Combine(tempPath, "JoinedDataPreview.html")
            IO.File.WriteAllText(htmlFile, htmlContent)

            ' Open in default browser
            OpenFileInBrowser(htmlFile)

        Catch ex As Exception
            MessageBox.Show($"Error displaying joined data: {ex.Message}",
                          "Display Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Shared Function GenerateJoinedDataHtml(joinedData As DataTable) As String
        Dim sb As New Text.StringBuilder()

        ' Add HTML header and styles
        sb.Append(GenerateHtmlHeader("Joined Data Preview"))

        ' Add summary
        sb.AppendLine("<div class='summary'>")
        sb.AppendLine($"<h2>Data Join Summary</h2>")
        sb.AppendLine($"<p>Total Records: {joinedData.Rows.Count}</p>")
        sb.AppendLine($"<p>Total Columns: {joinedData.Columns.Count}</p>")
        sb.AppendLine("</div>")

        ' Create table
        sb.AppendLine("<table>")

        ' Add headers
        sb.AppendLine("<tr>")
        For Each column As DataColumn In joinedData.Columns
            sb.AppendLine($"<th>{column.ColumnName}</th>")
        Next
        sb.AppendLine("</tr>")

        ' Add data rows
        For Each row As DataRow In joinedData.Rows
            sb.AppendLine("<tr>")
            For Each item In row.ItemArray
                sb.AppendLine($"<td>{item}</td>")
            Next
            sb.AppendLine("</tr>")
        Next

        sb.AppendLine("</table>")
        sb.Append(GenerateHtmlFooter(True))

        Return sb.ToString()
    End Function

    Public Shared Function VerifyAndDisplayExcelContent(excelPath As String, Optional showInBrowser As Boolean = True) As String
        Dim excel As Object = Nothing
        Dim workbook As Object = Nothing
        Dim worksheet As Object = Nothing
        Dim htmlContent As New Text.StringBuilder()

        Try
            ' Verify file exists
            If Not IO.File.Exists(excelPath) Then
                Throw New IO.FileNotFoundException("Excel file not found.", excelPath)
            End If

            ' Verify file extension
            Dim extension = IO.Path.GetExtension(excelPath).ToLower()
            If Not {".xls", ".xlsx", ".xlsm"}.Contains(extension) Then
                Throw New ArgumentException("Invalid file type. Please select an Excel file (.xls, .xlsx, or .xlsm).")
            End If

            ' Create Excel application
            excel = CreateObject("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            ' Open workbook
            workbook = excel.Workbooks.Open(excelPath)
            worksheet = workbook.Worksheets(1)
            Dim usedRange = worksheet.UsedRange
            Dim rowCount = usedRange.Rows.Count
            Dim colCount = usedRange.Columns.Count

            ' Start building HTML content
            htmlContent.AppendLine("<!DOCTYPE html>")
            htmlContent.AppendLine("<html lang='en'>")
            htmlContent.AppendLine("<head>")
            htmlContent.AppendLine("    <meta charset='UTF-8'>")
            htmlContent.AppendLine("    <meta name='viewport' content='width=device-width, initial-scale=1.0'>")
            htmlContent.AppendLine("    <title>Excel Content Preview</title>")
            htmlContent.AppendLine("    <style>")
            htmlContent.AppendLine("        body {")
            htmlContent.AppendLine("            font-family: Arial, sans-serif;")
            htmlContent.AppendLine("            margin: 20px;")
            htmlContent.AppendLine("            background-color: #f5f5f5;")
            htmlContent.AppendLine("        }")
            htmlContent.AppendLine("        .container {")
            htmlContent.AppendLine("            background-color: white;")
            htmlContent.AppendLine("            padding: 20px;")
            htmlContent.AppendLine("            border-radius: 8px;")
            htmlContent.AppendLine("            box-shadow: 0 2px 4px rgba(0,0,0,0.1);")
            htmlContent.AppendLine("        }")
            htmlContent.AppendLine("        h1 {")
            htmlContent.AppendLine("            color: #2c3e50;")
            htmlContent.AppendLine("            text-align: center;")
            htmlContent.AppendLine("            margin-bottom: 20px;")
            htmlContent.AppendLine("        }")
            htmlContent.AppendLine("        .excel-info {")
            htmlContent.AppendLine("            background-color: #f8f9fa;")
            htmlContent.AppendLine("            padding: 15px;")
            htmlContent.AppendLine("            border-radius: 5px;")
            htmlContent.AppendLine("            margin-bottom: 20px;")
            htmlContent.AppendLine("        }")
            htmlContent.AppendLine("        .table-container {")
            htmlContent.AppendLine("            overflow-x: auto;")
            htmlContent.AppendLine("        }")
            htmlContent.AppendLine("        table {")
            htmlContent.AppendLine("            width: 100%;")
            htmlContent.AppendLine("            border-collapse: collapse;")
            htmlContent.AppendLine("            margin-top: 10px;")
            htmlContent.AppendLine("        }")
            htmlContent.AppendLine("        th {")
            htmlContent.AppendLine("            background-color: #3498db;")
            htmlContent.AppendLine("            color: white;")
            htmlContent.AppendLine("            padding: 12px;")
            htmlContent.AppendLine("            text-align: left;")
            htmlContent.AppendLine("            position: sticky;")
            htmlContent.AppendLine("            top: 0;")
            htmlContent.AppendLine("        }")
            htmlContent.AppendLine("        td {")
            htmlContent.AppendLine("            padding: 10px;")
            htmlContent.AppendLine("            border: 1px solid #ddd;")
            htmlContent.AppendLine("        }")
            htmlContent.AppendLine("        tr:nth-child(even) {")
            htmlContent.AppendLine("            background-color: #f2f2f2;")
            htmlContent.AppendLine("        }")
            htmlContent.AppendLine("        tr:nth-child(odd) {")
            htmlContent.AppendLine("            background-color: #ffffff;")
            htmlContent.AppendLine("        }")
            htmlContent.AppendLine("        tr:hover {")
            htmlContent.AppendLine("            background-color: #e8f4f8;")
            htmlContent.AppendLine("        }")
            htmlContent.AppendLine("    </style>")
            htmlContent.AppendLine("</head>")
            htmlContent.AppendLine("<body>")
            htmlContent.AppendLine("    <div class='container'>")
            htmlContent.AppendLine("        <h1>Excel Content Preview</h1>")

            ' Add file information
            htmlContent.AppendLine("        <div class='excel-info'>")
            htmlContent.AppendLine("            <p><strong>File:</strong> " & IO.Path.GetFileName(excelPath) & "</p>")
            htmlContent.AppendLine("            <p><strong>Total Rows:</strong> " & rowCount & "</p>")
            htmlContent.AppendLine("            <p><strong>Total Columns:</strong> " & colCount & "</p>")
            htmlContent.AppendLine("            <p><strong>Last Updated:</strong> " & IO.File.GetLastWriteTime(excelPath) & "</p>")
            htmlContent.AppendLine("        </div>")

            ' Start table
            htmlContent.AppendLine("        <div class='table-container'>")
            htmlContent.AppendLine("            <table>")

            ' Headers
            htmlContent.AppendLine("                <tr>")
            For col = 1 To colCount
                Dim headerValue = worksheet.Cells(1, col).Text
                htmlContent.AppendLine("                    <th>" & Web.HttpUtility.HtmlEncode(headerValue) & "</th>")
            Next
            htmlContent.AppendLine("                </tr>")

            ' Data rows
            For row = 2 To rowCount
                htmlContent.AppendLine("                <tr>")
                For col = 1 To colCount
                    Dim cellValue = worksheet.Cells(row, col).Text
                    htmlContent.AppendLine("                    <td>" & Web.HttpUtility.HtmlEncode(cellValue) & "</td>")
                Next
                htmlContent.AppendLine("                </tr>")
            Next

            htmlContent.AppendLine("            </table>")
            htmlContent.AppendLine("        </div>")
            htmlContent.AppendLine("    </div>")
            htmlContent.AppendLine("</body>")
            htmlContent.AppendLine("</html>")

            ' If showInBrowser is true, display the content and return nothing
            If showInBrowser Then
                Dim tempPath As String = Environment.GetEnvironmentVariable("TEMP")
                Dim htmlFile As String = IO.Path.Combine(tempPath, "ExcelPreview_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".html")
                IO.File.WriteAllText(htmlFile, htmlContent.ToString())

                ' Use ProcessStartInfo to properly open the file in the default browser
                Dim psi As New ProcessStartInfo With {
                    .FileName = htmlFile,
                    .UseShellExecute = True,
                    .Verb = "open"
                }
                Process.Start(psi)
                Return Nothing
            End If

            ' If showInBrowser is false, return the HTML content
            Return htmlContent.ToString()

        Catch ex As Exception
            MessageBox.Show("Error verifying Excel file: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        Finally
            ' Cleanup
            If worksheet IsNot Nothing Then ReleaseComObject(worksheet)
            If workbook IsNot Nothing Then
                workbook.Close(False)
                ReleaseComObject(workbook)
            End If
            If excel IsNot Nothing Then
                excel.Quit()
                ReleaseComObject(excel)
            End If
        End Try
    End Function
End Class