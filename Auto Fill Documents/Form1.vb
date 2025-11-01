Public Class Frm_MailMerge
    ' Flag to prevent recursive checkbox events
    Private isProcessingCheckbox As Boolean = False
    Private Const DEFAULT_FOLDER As String = "AUTOFILL_DOCS"

    Private progressBar As ProgressBar
    Private progressLabel As Label
    Private WithEvents ControlsHelpToolStripMenuItem As ToolStripMenuItem
    Private WithEvents DataVerificationToolStripMenuItem As ToolStripMenuItem

    ' Add at the top of the class
    Private Sub HandleError(ex As Exception, operation As String)
        Try
            ' Log the error (you can implement logging if needed)
            Debug.WriteLine($"Error in {operation}: {ex.Message}")
            Debug.WriteLine($"Stack Trace: {ex.StackTrace}")

            ' Show user-friendly message based on error type
            Dim message As String = "An error occurred during " & operation & "."
            Dim title As String = "Operation Error"

            Select Case True
                Case TypeOf ex Is UnauthorizedAccessException
                    message = "Access denied. Please check your permissions and try again."
                    title = "Access Error"

                Case TypeOf ex Is IO.IOException
                    message = "File access error. The file might be in use or locked."
                    title = "File Access Error"

                Case TypeOf ex Is System.Runtime.InteropServices.COMException
                    message = "Error communicating with Microsoft Office. Please ensure Excel and Word are properly installed."
                    title = "Office Communication Error"

                Case TypeOf ex Is OutOfMemoryException
                    message = "Not enough memory to complete the operation. Please close some applications and try again."
                    title = "Memory Error"

                Case TypeOf ex Is ArgumentException
                    message = "Invalid input or argument. Please check your selections and try again."
                    title = "Input Error"

                Case Else
                    message &= vbCrLf & "Details: " & ex.Message
            End Select

            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Error)

        Catch innerEx As Exception
            ' Fallback error handling if even the error handler fails
            MessageBox.Show("A critical error occurred. The application will try to continue.",
                           "Critical Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' Always reset progress
            UpdateProgress(0, "")
        End Try
    End Sub

    ' Add global exception handler to Form_Load
    Protected Overrides Sub OnLoad(e As EventArgs)
        Try
            MyBase.OnLoad(e)
            ' Set up global exception handling
            AddHandler Application.ThreadException, AddressOf HandleThreadException
            AddHandler AppDomain.CurrentDomain.UnhandledException, AddressOf HandleUnhandledException

            Form1_Load(Me, e)
        Catch ex As Exception
            HandleError(ex, "application initialization")
        End Try
    End Sub

    Private Sub HandleThreadException(sender As Object, e As Threading.ThreadExceptionEventArgs)
        HandleError(e.Exception, "unexpected thread operation")
    End Sub

    Private Sub HandleUnhandledException(sender As Object, e As UnhandledExceptionEventArgs)
        Try
            If e.ExceptionObject IsNot Nothing AndAlso TypeOf e.ExceptionObject Is Exception Then
                HandleError(DirectCast(e.ExceptionObject, Exception), "unexpected operation")
            End If
        Catch ex As Exception
            MessageBox.Show("A critical error occurred. The application will try to continue.",
                           "Critical Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Add safe COM object cleanup
    Private Function SafeReleaseCom(obj As Object) As Boolean
        Try
            If obj IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
                Return True
            End If
        Catch ex As Exception
            Debug.WriteLine($"Error releasing COM object: {ex.Message}")
        Finally
            obj = Nothing
        End Try
        Return False
    End Function

    ' Add safe file operation wrapper
    Private Function SafeFileOperation(operation As Action, operationName As String) As Boolean
        Try
            operation.Invoke()
            Return True
        Catch ex As UnauthorizedAccessException
            HandleError(ex, operationName)
        Catch ex As IO.IOException
            HandleError(ex, operationName)
        Catch ex As Exception
            HandleError(ex, operationName)
        End Try
        Return False
    End Function

    ' Add safe process termination
    Private Sub SafeKillProcess(processName As String)
        Try
            Dim processes As Process() = Process.GetProcessesByName(processName)
            If processes IsNot Nothing AndAlso processes.Length > 0 Then
                For Each proc As Process In processes
                    Try
                        If proc IsNot Nothing Then
                            If Not proc.HasExited Then
                                proc.Kill()
                                proc.WaitForExit(1000) ' Wait up to 1 second
                            End If
                            proc.Dispose()
                        End If
                    Catch ex As Exception
                        Debug.WriteLine($"Error killing process {processName}: {ex.Message}")
                    End Try
                Next
            End If
        Catch ex As Exception
            Debug.WriteLine($"Error in SafeKillProcess: {ex.Message}")
        End Try
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try

            ' Initialize ComboBoxes
            Reusable_Class_Functions.InitializeComboBoxes(
                ComboBox1, ComboBox2, ComboBox3, ComboBox4,
                ComboBox5)

            ' Initialize states for controls with checked ignore boxes
            InitializeControlStates()

            ' Set default output path
            TextBox6.Text = Reusable_Class_Functions.GetDefaultOutputPath()
        Catch ex As Exception
            MessageBox.Show("Error initializing form: " & ex.Message, "Initialization Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub UpdateProgress(percentage As Integer, status As String)
        Reusable_Class_Functions.UpdateProgressBar(ProgressBar1, Label7, percentage, status)
    End Sub

    Private Sub ReleaseComObject(obj As Object)
        Reusable_Class_Functions.ReleaseComObject(obj)
    End Sub

    Private Sub UpdateGroupBox4State()
        ' Count how many Excel files are active (not ignored)
        Dim activeExcelFiles = 0
        If Not CheckBox3.Checked AndAlso Not String.IsNullOrEmpty(TextBox3.Text) Then activeExcelFiles += 1
        If Not CheckBox4.Checked AndAlso Not String.IsNullOrEmpty(TextBox4.Text) Then activeExcelFiles += 1
        If Not CheckBox5.Checked AndAlso Not String.IsNullOrEmpty(TextBox5.Text) Then activeExcelFiles += 1

        ' Enable GroupBox4 only if there are more than one active Excel file
        GroupBox4.Enabled = (activeExcelFiles > 1)

        ' If disabling, clear all ComboBoxes in GroupBox4

    End Sub

    Private Sub InitializeControlStates()
        ' Initialize GroupBox1 controls
        Button1.Enabled = Not CheckBox1.Checked
        Button2.Enabled = Not CheckBox2.Checked
        If CheckBox1.Checked Then TextBox1.Text = ""
        If CheckBox2.Checked Then TextBox2.Text = ""

        ' Initialize GroupBox2 controls
        Button3.Enabled = Not CheckBox3.Checked
        Button4.Enabled = Not CheckBox4.Checked
        Button5.Enabled = Not CheckBox5.Checked
        If CheckBox3.Checked Then TextBox3.Text = ""
        If CheckBox4.Checked Then TextBox4.Text = ""
        If CheckBox5.Checked Then TextBox5.Text = ""

        ' Update GroupBox4 state
        UpdateGroupBox4State()
    End Sub

    Private Function AreAllCheckedInGroup1() As Boolean
        Return CheckBox1.Checked AndAlso CheckBox2.Checked
    End Function

    Private Function AreAllCheckedInGroup2() As Boolean
        Return CheckBox3.Checked AndAlso CheckBox4.Checked AndAlso CheckBox5.Checked
    End Function

    Private Sub ShowErrorMessage()
        MessageBox.Show("At least one document/file must not be ignored in each group.", "Invalid Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    End Sub

    Private Sub HandleCheckBoxChange(checkBox As CheckBox, button As Button, textBox As TextBox, groupNumber As Integer)
        If isProcessingCheckbox Then Return

        Try
            isProcessingCheckbox = True

            Dim groupCheckBoxes As CheckBox()
            If groupNumber = 1 Then
                groupCheckBoxes = {CheckBox1, CheckBox2}
            Else
                groupCheckBoxes = {CheckBox3, CheckBox4, CheckBox5}
            End If

            Reusable_Class_Functions.HandleCheckBoxChange(
                checkBox, button, textBox, groupCheckBoxes, groupNumber,
                Sub() UpdateGroupBox4State())


        Finally
            isProcessingCheckbox = False
        End Try
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        HandleCheckBoxChange(CheckBox1, Button1, TextBox1, 1)
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        HandleCheckBoxChange(CheckBox2, Button2, TextBox2, 1)
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        HandleCheckBoxChange(CheckBox3, Button3, TextBox3, 2)
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        HandleCheckBoxChange(CheckBox4, Button4, TextBox4, 2)
    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        HandleCheckBoxChange(CheckBox5, Button5, TextBox5, 2)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            UpdateProgress(0, "Selecting Word document...")

            ' First show the file dialog to select the Word document
            Reusable_Class_Functions.ShowFileDialog(
                "Word Documents|*.doc;*.docx",
                "Select Word Document",
                TextBox1,
                {CheckBox1, CheckBox2},
                {TextBox1, TextBox2})

            ' If a file was selected (TextBox1 is not empty)
            If Not String.IsNullOrEmpty(TextBox1.Text) Then
                UpdateProgress(20, "Clearing existing field mappings...")
                ' Clear existing items in ComboBox1
                ComboBox1.Items.Clear()

                UpdateProgress(40, "Scanning document for merge fields...")
                ' Extract field names from the Word document
                Dim fieldNames = ExtractFieldNamesFromWord(TextBox1.Text)

                ' Check if any fields were found
                If fieldNames.Count > 0 Then
                    UpdateProgress(60, "Adding merge fields to dropdown...")
                    ' Add field names to ComboBox1
                    ComboBox1.Items.AddRange(fieldNames.ToArray())

                    UpdateProgress(80, "Setting default selection...")
                    ' Set the first item as default selection
                    ComboBox1.SelectedIndex = 0

                    UpdateProgress(100, "Document processed successfully!")
                Else
                    UpdateProgress(100, "No merge fields found!")
                    ' No fields found - show error message
                    MessageBox.Show("No merge fields (<<FieldName>>) found in the selected document. " &
                                  "Please revise the document to include merge fields or select another document.",
                                  "Invalid Document", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    ' Clear the TextBox1
                    TextBox1.Clear()
                End If
            End If

        Catch ex As Exception
            UpdateProgress(100, "Error processing document!")
            MessageBox.Show("Error processing Word document: " & ex.Message,
                          "Document Processing Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox1.Clear()
            ComboBox1.Items.Clear()
        Finally
            ' Reset progress bar after a delay
            System.Threading.Thread.Sleep(1000)
            UpdateProgress(0, "")
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            UpdateProgress(0, "Selecting Word document...")

            ' First show the file dialog to select the Word document
            Reusable_Class_Functions.ShowFileDialog(
                "Word Documents|*.doc;*.docx",
                "Select Word Document",
                TextBox2,
                {CheckBox1, CheckBox2},
                {TextBox1, TextBox2})

            ' If a file was selected (TextBox2 is not empty)
            If Not String.IsNullOrEmpty(TextBox2.Text) Then
                UpdateProgress(20, "Clearing existing field mappings...")
                ' Clear existing items in ComboBox2
                ComboBox2.Items.Clear()

                UpdateProgress(40, "Scanning document for merge fields...")
                ' Extract field names from the Word document
                Dim fieldNames = ExtractFieldNamesFromWord(TextBox2.Text)

                ' Check if any fields were found
                If fieldNames.Count > 0 Then
                    UpdateProgress(60, "Adding merge fields to dropdown...")
                    ' Add field names to ComboBox2
                    ComboBox2.Items.AddRange(fieldNames.ToArray())

                    UpdateProgress(80, "Setting default selection...")
                    ' Set the first item as default selection
                    ComboBox2.SelectedIndex = 0

                    UpdateProgress(100, "Document processed successfully!")
                Else
                    UpdateProgress(100, "No merge fields found!")
                    ' No fields found - show error message
                    MessageBox.Show("No merge fields (<<FieldName>>) found in the selected document. " &
                                  "Please revise the document to include merge fields or select another document.",
                                  "Invalid Document", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    ' Clear the TextBox2
                    TextBox2.Clear()
                End If
            End If

        Catch ex As Exception
            UpdateProgress(100, "Error processing document!")
            MessageBox.Show("Error processing Word document: " & ex.Message,
                          "Document Processing Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox2.Clear()
            ComboBox2.Items.Clear()
        Finally
            ' Reset progress bar after a delay
            System.Threading.Thread.Sleep(1000)
            UpdateProgress(0, "")
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            UpdateProgress(0, "Selecting Excel file...")

            ' First show the file dialog to select the Excel file
            Reusable_Class_Functions.ShowFileDialog(
                "Excel Files|*.xls;*.xlsx;*.xlsm",
                "Select Excel File",
                TextBox3,
                {CheckBox3, CheckBox4, CheckBox5},
                {TextBox3, TextBox4, TextBox5},
                Sub() UpdateGroupBox4State())

            ' If a file was selected (TextBox3 is not empty)
            If Not String.IsNullOrEmpty(TextBox3.Text) Then
                UpdateProgress(20, "Clearing existing field mappings...")
                ' Clear existing items in ComboBox3
                ComboBox3.Items.Clear()

                UpdateProgress(40, "Reading Excel headers...")
                ' Get Excel headers
                Dim headers = GetExcelHeaders(TextBox3.Text)

                ' Check if any headers were found
                If headers.Count > 0 Then
                    UpdateProgress(60, "Adding headers to dropdown...")
                    ' Add headers to ComboBox3
                    ComboBox3.Items.AddRange(headers.ToArray())

                    UpdateProgress(80, "Setting default selection...")
                    ' Set the first item as default selection
                    ComboBox3.SelectedIndex = 0

                    UpdateProgress(100, "Excel file processed successfully!")
                Else
                    UpdateProgress(100, "No headers found!")
                    ' No headers found - show error message
                    MessageBox.Show("No headers found in the first row of the Excel file. " &
                                  "Please ensure the first row contains field names or select another file.",
                                  "Invalid Excel File", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    ' Clear the TextBox3
                    TextBox3.Clear()
                End If
            End If

        Catch ex As Exception
            UpdateProgress(100, "Error processing Excel file!")
            MessageBox.Show("Error processing Excel file: " & ex.Message,
                          "Excel Processing Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox3.Clear()
            ComboBox3.Items.Clear()
        Finally
            ' Reset progress bar after a delay
            System.Threading.Thread.Sleep(1000)
            UpdateProgress(0, "")
        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            UpdateProgress(0, "Selecting Excel file...")

            ' First show the file dialog to select the Excel file
            Reusable_Class_Functions.ShowFileDialog(
                "Excel Files|*.xls;*.xlsx;*.xlsm",
                "Select Excel File",
                TextBox4,
                {CheckBox3, CheckBox4, CheckBox5},
                {TextBox3, TextBox4, TextBox5},
                Sub() UpdateGroupBox4State())

            ' If a file was selected (TextBox4 is not empty)
            If Not String.IsNullOrEmpty(TextBox4.Text) Then
                UpdateProgress(20, "Clearing existing field mappings...")
                ' Clear existing items in ComboBox4
                ComboBox4.Items.Clear()

                UpdateProgress(40, "Reading Excel headers...")
                ' Get Excel headers
                Dim headers = GetExcelHeaders(TextBox4.Text)

                ' Check if any headers were found
                If headers.Count > 0 Then
                    UpdateProgress(60, "Adding headers to dropdown...")
                    ' Add headers to ComboBox4
                    ComboBox4.Items.AddRange(headers.ToArray())

                    UpdateProgress(80, "Setting default selection...")
                    ' Set the first item as default selection
                    ComboBox4.SelectedIndex = 0

                    UpdateProgress(100, "Excel file processed successfully!")
                Else
                    UpdateProgress(100, "No headers found!")
                    ' No headers found - show error message
                    MessageBox.Show("No headers found in the first row of the Excel file. " &
                                  "Please ensure the first row contains field names or select another file.",
                                  "Invalid Excel File", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    ' Clear the TextBox4
                    TextBox4.Clear()
                End If
            End If

        Catch ex As Exception
            UpdateProgress(100, "Error processing Excel file!")
            MessageBox.Show("Error processing Excel file: " & ex.Message,
                          "Excel Processing Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox4.Clear()
            ComboBox4.Items.Clear()
        Finally
            ' Reset progress bar after a delay
            System.Threading.Thread.Sleep(1000)
            UpdateProgress(0, "")
        End Try
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Try
            UpdateProgress(0, "Selecting Excel file...")

            ' First show the file dialog to select the Excel file
            Reusable_Class_Functions.ShowFileDialog(
                "Excel Files|*.xls;*.xlsx;*.xlsm",
                "Select Excel File",
                TextBox5,
                {CheckBox3, CheckBox4, CheckBox5},
                {TextBox3, TextBox4, TextBox5},
                Sub() UpdateGroupBox4State())

            ' If a file was selected (TextBox5 is not empty)
            If Not String.IsNullOrEmpty(TextBox5.Text) Then
                UpdateProgress(20, "Clearing existing field mappings...")
                ' Clear existing items in ComboBox5
                ComboBox5.Items.Clear()

                UpdateProgress(40, "Reading Excel headers...")
                ' Get Excel headers
                Dim headers = GetExcelHeaders(TextBox5.Text)

                ' Check if any headers were found
                If headers.Count > 0 Then
                    UpdateProgress(60, "Adding headers to dropdown...")
                    ' Add headers to ComboBox5
                    ComboBox5.Items.AddRange(headers.ToArray())

                    UpdateProgress(80, "Setting default selection...")
                    ' Set the first item as default selection
                    ComboBox5.SelectedIndex = 0

                    UpdateProgress(100, "Excel file processed successfully!")
                Else
                    UpdateProgress(100, "No headers found!")
                    ' No headers found - show error message
                    MessageBox.Show("No headers found in the first row of the Excel file. " &
                                  "Please ensure the first row contains field names or select another file.",
                                  "Invalid Excel File", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    ' Clear the TextBox5
                    TextBox5.Clear()
                End If
            End If

        Catch ex As Exception
            UpdateProgress(100, "Error processing Excel file!")
            MessageBox.Show("Error processing Excel file: " & ex.Message,
                          "Excel Processing Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox5.Clear()
            ComboBox5.Items.Clear()
        Finally
            ' Reset progress bar after a delay
            System.Threading.Thread.Sleep(1000)
            UpdateProgress(0, "")
        End Try
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Try
            Using folderDialog As New FolderBrowserDialog
                folderDialog.Description = "Select Output Folder"
                folderDialog.UseDescriptionForTitle = True

                If folderDialog.ShowDialog = DialogResult.OK Then
                    TextBox6.Text = folderDialog.SelectedPath
                    ' Check if the selected path is different from the default path
                    Dim defaultPath = IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), DEFAULT_FOLDER)
                    Dim selectedPath = folderDialog.SelectedPath.TrimEnd(IO.Path.DirectorySeparatorChar)
                    defaultPath = defaultPath.TrimEnd(IO.Path.DirectorySeparatorChar)

                    ' Also check for D: drive default path
                    Dim dDrivePath = ""
                    Dim dDrive = IO.DriveInfo.GetDrives.FirstOrDefault(Function(d) d.Name.StartsWith("D"))
                    If dDrive IsNot Nothing AndAlso dDrive.IsReady Then
                        dDrivePath = IO.Path.Combine(dDrive.RootDirectory.FullName, DEFAULT_FOLDER).TrimEnd(IO.Path.DirectorySeparatorChar)
                    End If

                    ' Set CheckBox6 to unchecked if the selected path is different from both default paths
                    CheckBox6.Checked = String.Equals(selectedPath, defaultPath, StringComparison.OrdinalIgnoreCase) OrElse
                                      Not String.IsNullOrEmpty(dDrivePath) AndAlso String.Equals(selectedPath, dDrivePath, StringComparison.OrdinalIgnoreCase)
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show("Error selecting folder: " & ex.Message, "Folder Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged
        If CheckBox6.Checked Then
            ' When checked, restore default path
            TextBox6.Text = Reusable_Class_Functions.GetDefaultOutputPath()
        End If
    End Sub

    Private Function ExtractFieldNamesFromWord(filePath As String) As List(Of String)
        Try
            Return Reusable_Class_Functions.ExtractFieldNamesFromWord(filePath)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Word Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return New List(Of String)
        End Try
    End Function

    Private Function ShowFieldMatchingDialog(wordFields As List(Of String), excelHeaders As List(Of String), documentNumber As Integer) As Dictionary(Of String, String)
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
        panel.RowCount = wordFields.Count + 2 ' +1 for header, +1 for button
        panel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))
        panel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))
        panel.Padding = New Padding(10)

        ' Add headers
        Dim headerLabel1 As New Label()
        headerLabel1.Text = $"Word Fields (Document {documentNumber})"
        headerLabel1.Font = New Font(headerLabel1.Font, FontStyle.Bold)
        panel.Controls.Add(headerLabel1, 0, 0)

        Dim headerLabel2 As New Label()
        headerLabel2.Text = "Excel Headers"
        headerLabel2.Font = New Font(headerLabel2.Font, FontStyle.Bold)
        panel.Controls.Add(headerLabel2, 1, 0)

        ' Dictionary to store ComboBoxes
        Dim fieldMatches As New Dictionary(Of String, ComboBox)

        ' Add field matching controls
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
            combo.Items.Add("(None)") ' Add option for no match
            combo.SelectedItem = "(None)"
            panel.Controls.Add(combo, 1, i + 1)

            fieldMatches.Add(wordFields(i), combo)
        Next

        ' Add OK button
        Dim btnOK As New Button()
        btnOK.Text = "OK"
        btnOK.DialogResult = DialogResult.OK
        btnOK.Dock = DockStyle.Bottom
        panel.Controls.Add(btnOK, 0, wordFields.Count + 1)
        panel.SetColumnSpan(btnOK, 2)

        matchForm.Controls.Add(panel)

        ' Show dialog and process results
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

    Private Sub ProcessMailMerge(excelPath As String)
        Dim excel As Object = Nothing
        Dim workbook As Object = Nothing
        Dim worksheet As Object = Nothing
        Dim usedRange As Object = Nothing
        Dim word As Object = Nothing

        Try
            ' Create timestamped subfolder for this batch
            Dim timestamp As String = DateTime.Now.ToString("yyyyMMdd_HHmmss")
            Dim batchFolder As String = IO.Path.Combine(TextBox6.Text, timestamp)

            ' Create the subfolder
            If Not IO.Directory.Exists(batchFolder) Then
                IO.Directory.CreateDirectory(batchFolder)
            End If

            ' Reset progress bar
            ProgressBar1.Value = 0
            UpdateProgress(0, "Initializing...")

            ' Create Excel Application instance
            excel = CreateObject("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            UpdateProgress(10, "Opening Excel file...")

            ' Open Excel workbook
            workbook = excel.Workbooks.Open(excelPath)
            worksheet = workbook.Worksheets(1)
            usedRange = worksheet.UsedRange
            UpdateProgress(20, "Reading Excel data...")

            If usedRange.Rows.Count < 2 Then
                MessageBox.Show("Excel file must have at least one data row.", "No Data", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' Get column headers
            Dim headers = GetExcelHeaders(excelPath)
            If headers.Count = 0 Then Return
            UpdateProgress(30, "Creating Word instance...")

            ' Create Word Application instance
            word = CreateObject("Word.Application")
            word.Visible = False

            ' Get Word fields and check for matches
            Dim doc1Fields As List(Of String) = Nothing
            Dim doc2Fields As List(Of String) = Nothing
            Dim fieldMappings1 As Dictionary(Of String, String) = Nothing
            Dim fieldMappings2 As Dictionary(Of String, String) = Nothing

            UpdateProgress(40, "Checking Document 1...")
            If Not CheckBox1.Checked AndAlso Not String.IsNullOrEmpty(TextBox1.Text) Then
                doc1Fields = ExtractFieldNamesFromWord(TextBox1.Text)
                If Not doc1Fields.All(Function(f) headers.Contains(f)) Then
                    MessageBox.Show("Some fields in Document 1 don't match Excel headers. Please match them manually.", "Field Matching Required", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    fieldMappings1 = ShowFieldMatchingDialog(doc1Fields, headers, 1)
                    If fieldMappings1.Count = 0 Then Return
                End If
            End If

            UpdateProgress(50, "Checking Document 2...")
            If Not CheckBox2.Checked AndAlso Not String.IsNullOrEmpty(TextBox2.Text) Then
                doc2Fields = ExtractFieldNamesFromWord(TextBox2.Text)
                If Not doc2Fields.All(Function(f) headers.Contains(f)) Then
                    MessageBox.Show("Some fields in Document 2 don't match Excel headers. Please match them manually.", "Field Matching Required", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    fieldMappings2 = ShowFieldMatchingDialog(doc2Fields, headers, 2)
                    If fieldMappings2.Count = 0 Then Return
                End If
            End If

            ' Process based on RadioButton selection
            If RadioButton1.Checked Then
                ' Generate separate documents for each row
                Dim rowCount = usedRange.Rows.Count
                For rowIndex = 2 To rowCount
                    Dim progress = 50 + (rowIndex - 2) * 50 \ (rowCount - 1)
                    UpdateProgress(progress, $"Processing row {rowIndex - 1} of {rowCount - 1}...")

                    ' Process each Word document for this row
                    If Not CheckBox1.Checked AndAlso Not String.IsNullOrEmpty(TextBox1.Text) Then
                        Reusable_Class_Functions.ProcessWordDocument(TextBox1.Text, word, worksheet, rowIndex, headers, "Doc1", batchFolder, fieldMappings1)
                    End If
                    If Not CheckBox2.Checked AndAlso Not String.IsNullOrEmpty(TextBox2.Text) Then
                        Reusable_Class_Functions.ProcessWordDocument(TextBox2.Text, word, worksheet, rowIndex, headers, "Doc2", batchFolder, fieldMappings2)
                    End If
                Next
            Else
                ' Generate single documents with page breaks
                If Not CheckBox1.Checked AndAlso Not String.IsNullOrEmpty(TextBox1.Text) Then
                    ProcessMultiPageDocument(TextBox1.Text, word, worksheet, usedRange.Rows.Count, headers, "Doc1", batchFolder, fieldMappings1)
                End If
                If Not CheckBox2.Checked AndAlso Not String.IsNullOrEmpty(TextBox2.Text) Then
                    ProcessMultiPageDocument(TextBox2.Text, word, worksheet, usedRange.Rows.Count, headers, "Doc2", batchFolder, fieldMappings2)
                End If
            End If

            UpdateProgress(100, "Completed!")
            MessageBox.Show($"Documents have been generated successfully in folder: {timestamp}", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show("Error during mail merge: " & ex.Message, "Mail Merge Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' Reset progress bar
            ProgressBar1.Value = 0
            If Label7 IsNot Nothing Then
                Label7.Text = ""
            End If

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
            If word IsNot Nothing Then
                word.Quit()
                ReleaseComObject(word)
            End If
        End Try
    End Sub

    Private Sub ProcessMultiPageDocument(wordPath As String, word As Object, worksheet As Object, rowCount As Integer, headers As List(Of String), docPrefix As String, outputFolder As String, Optional fieldMappings As Dictionary(Of String, String) = Nothing)
        Dim sourceDoc As Object = Nothing
        Dim targetDoc As Object = Nothing
        Dim selection As Object = Nothing

        Try
            ' Create new document first
            targetDoc = word.Documents.Add()

            ' Open source document
            sourceDoc = word.Documents.Open(wordPath)

            ' Copy content from source to target
            sourceDoc.Content.Copy()
            targetDoc.Content.Paste()

            ' Process first row
            ProcessRowInDocument(targetDoc, worksheet, 2, headers, fieldMappings)

            ' For remaining rows, copy the template and process
            For rowIndex = 3 To rowCount
                Dim progress = 50 + (rowIndex - 2) * 50 \ (rowCount - 1)
                UpdateProgress(progress, $"Processing row {rowIndex - 1} of {rowCount - 1}...")

                ' Insert page break
                selection = targetDoc.Range(targetDoc.Content.End - 1, targetDoc.Content.End - 1)
                selection.InsertBreak(7) ' wdPageBreak = 7

                ' Copy template content
                sourceDoc.Content.Copy()
                selection = targetDoc.Range(targetDoc.Content.End - 1, targetDoc.Content.End - 1)
                selection.Paste()

                ' Process the newly added content
                ProcessRowInDocument(targetDoc, worksheet, rowIndex, headers, fieldMappings)
            Next

            ' Get original file extension
            Dim originalExtension As String = IO.Path.GetExtension(wordPath)

            ' Generate output filename for the combined document
            Dim outputFileName = $"{docPrefix}_Combined{originalExtension}"
            Dim outputPath = IO.Path.Combine(outputFolder, outputFileName)

            ' Save the combined document
            targetDoc.SaveAs(outputPath)

        Finally
            If selection IsNot Nothing Then ReleaseComObject(selection)
            If sourceDoc IsNot Nothing Then
                sourceDoc.Close(False)
                ReleaseComObject(sourceDoc)
            End If
            If targetDoc IsNot Nothing Then
                targetDoc.Close(False)
                ReleaseComObject(targetDoc)
            End If
        End Try
    End Sub

    Private Sub ProcessRowInDocument(doc As Object, worksheet As Object, rowIndex As Integer, headers As List(Of String), fieldMappings As Dictionary(Of String, String))
        Dim range As Object = Nothing
        Try
            ' Get the last page range
            Dim lastParagraph As Object = doc.Paragraphs.Last
            Dim startPos As Long = lastParagraph.Range.Start
            range = doc.Range(startPos, doc.Content.End)

            If fieldMappings IsNot Nothing Then
                ' Use manual field mappings
                For Each mapping In fieldMappings
                    Dim columnIndex = headers.IndexOf(mapping.Value) + 1
                    If columnIndex > 0 Then
                        Dim cellValue = worksheet.Cells(rowIndex, columnIndex).Value
                        Dim replacementText = If(cellValue IsNot Nothing, CStr(cellValue), "")
                        ExecuteFindAndReplace(range.Find, mapping.Key, replacementText)
                    End If
                Next
            Else
                ' Use direct field matching
                For columnIndex = 1 To headers.Count
                    Dim fieldName = headers(columnIndex - 1)
                    Dim cellValue = worksheet.Cells(rowIndex, columnIndex).Value
                    Dim replacementText = If(cellValue IsNot Nothing, CStr(cellValue), "")
                    ExecuteFindAndReplace(range.Find, fieldName, replacementText)
                Next
            End If
        Finally
            If range IsNot Nothing Then ReleaseComObject(range)
        End Try
    End Sub

    Private Sub ExecuteFindAndReplace(findObject As Object, fieldName As String, replacementText As String)
        With findObject
            .ClearFormatting()
            .Text = $"<<{fieldName}>>"
            .Replacement.ClearFormatting()
            .Replacement.Text = replacementText
            .Forward = True
            .Wrap = 1 ' wdFindStop
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute(Replace:=2) ' wdReplaceAll
        End With
    End Sub

    Private Function GetExcelHeaders(excelPath As String) As List(Of String)
        Try
            Return Reusable_Class_Functions.GetExcelHeaders(excelPath)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Excel Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return New List(Of String)
        End Try
    End Function

    Private Sub ShowAboutInfo()
        Try
            Dim tempPath As String = Environment.GetEnvironmentVariable("TEMP")
            Dim htmlFile As String = IO.Path.Combine(tempPath, "AutoFillDocsSynopsis.html")

            ' Check if file exists and is not older than 24 hours
            Dim shouldGenerateFile As Boolean = True
            If IO.File.Exists(htmlFile) Then
                Dim fileInfo As New IO.FileInfo(htmlFile)
                If DateTime.Now.Subtract(fileInfo.LastWriteTime).TotalHours < 24 Then
                    shouldGenerateFile = False
                End If
            End If

            If shouldGenerateFile Then
                ' Generate HTML content
                Dim htmlContent As String = GenerateAboutHtml()
                ' Save to temp file
                IO.File.WriteAllText(htmlFile, htmlContent, System.Text.Encoding.UTF8)
            End If

            ' Open in browser
            Reusable_Class_Functions.OpenFileInBrowser(htmlFile)

        Catch ex As Exception
            MessageBox.Show($"Error showing about information: {ex.Message}", "About Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function GenerateAboutHtml() As String
        Dim sb As New Text.StringBuilder()

        sb.AppendLine("<!DOCTYPE html>")
        sb.AppendLine("<html lang='en'>")
        sb.AppendLine("<head>")
        sb.AppendLine("    <meta charset='UTF-8'>")
        sb.AppendLine("    <meta name='viewport' content='width=device-width, initial-scale=1.0'>")
        sb.AppendLine("    <title>Auto Fill Documents - Synopsis</title>")
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
        sb.AppendLine("        ul {")
        sb.AppendLine("            padding-left: 20px;")
        sb.AppendLine("        }")
        sb.AppendLine("        li {")
        sb.AppendLine("            margin-bottom: 8px;")
        sb.AppendLine("        }")
        sb.AppendLine("        .feature-section {")
        sb.AppendLine("            margin: 15px 0;")
        sb.AppendLine("            padding: 15px;")
        sb.AppendLine("            background-color: #f8f9fa;")
        sb.AppendLine("            border-radius: 5px;")
        sb.AppendLine("            border-left: 4px solid #3498db;")
        sb.AppendLine("        }")
        sb.AppendLine("        .footer {")
        sb.AppendLine("            text-align: center;")
        sb.AppendLine("            margin-top: 30px;")
        sb.AppendLine("            color: #7f8c8d;")
        sb.AppendLine("            font-size: 0.9em;")
        sb.AppendLine("        }")
        sb.AppendLine("    </style>")
        sb.AppendLine("</head>")
        sb.AppendLine("<body>")
        sb.AppendLine("    <div class='container'>")
        sb.AppendLine("        <h1>Auto Fill Documents - Application Synopsis</h1>")

        ' Core Features
        sb.AppendLine("        <div class='feature-section'>")
        sb.AppendLine("            <h2>1. Core Features</h2>")
        sb.AppendLine("            <ul>")
        sb.AppendLine("                <li>Advanced document automation tool for mail merge operations</li>")
        sb.AppendLine("                <li>Support for multiple Word documents and Excel data sources</li>")
        sb.AppendLine("                <li>Intelligent field mapping with automatic detection</li>")
        sb.AppendLine("                <li>Progress tracking with detailed status updates</li>")
        sb.AppendLine("                <li>Comprehensive error handling and recovery</li>")
        sb.AppendLine("            </ul>")
        sb.AppendLine("        </div>")

        ' Document Management
        sb.AppendLine("        <div class='feature-section'>")
        sb.AppendLine("            <h2>2. Document Management</h2>")
        sb.AppendLine("            <ul>")
        sb.AppendLine("                <li>Support for up to 2 Word templates with independent field mapping</li>")
        sb.AppendLine("                <li>Handles up to 3 Excel data sources with join capabilities</li>")
        sb.AppendLine("                <li>Selective processing with ignore options for each document</li>")
        sb.AppendLine("                <li>Automatic field detection in Word (<<FieldName>> format)</li>")
        sb.AppendLine("                <li>Excel header extraction and mapping</li>")
        sb.AppendLine("            </ul>")
        sb.AppendLine("        </div>")

        ' Data Processing
        sb.AppendLine("        <div class='feature-section'>")
        sb.AppendLine("            <h2>3. Data Processing</h2>")
        sb.AppendLine("            <ul>")
        sb.AppendLine("                <li>Advanced Excel file joining with field mapping</li>")
        sb.AppendLine("                <li>Data type preservation during processing</li>")
        sb.AppendLine("                <li>Support for multiple output formats:</li>")
        sb.AppendLine("                <ul>")
        sb.AppendLine("                    <li>Individual documents per data row</li>")
        sb.AppendLine("                    <li>Combined document with page breaks</li>")
        sb.AppendLine("                </ul>")
        sb.AppendLine("                <li>Automatic field matching with manual override options</li>")
        sb.AppendLine("            </ul>")
        sb.AppendLine("        </div>")

        ' User Interface
        sb.AppendLine("        <div class='feature-section'>")
        sb.AppendLine("            <h2>4. User Interface</h2>")
        sb.AppendLine("            <ul>")
        sb.AppendLine("                <li>Intuitive document selection and field mapping</li>")
        sb.AppendLine("                <li>Real-time progress tracking with status messages</li>")
        sb.AppendLine("                <li>Interactive file selection dialogs</li>")
        sb.AppendLine("                <li>Customizable output folder management</li>")
        sb.AppendLine("                <li>Quick access to output locations</li>")
        sb.AppendLine("            </ul>")
        sb.AppendLine("        </div>")

        ' Safety Features
        sb.AppendLine("        <div class='feature-section'>")
        sb.AppendLine("            <h2>5. Safety Features</h2>")
        sb.AppendLine("            <ul>")
        sb.AppendLine("                <li>Comprehensive error handling and recovery</li>")
        sb.AppendLine("                <li>Automatic cleanup of temporary files</li>")
        sb.AppendLine("                <li>Safe COM object management</li>")
        sb.AppendLine("                <li>Duplicate file selection prevention</li>")
        sb.AppendLine("                <li>Data validation before processing</li>")
        sb.AppendLine("                <li>Confirmation dialogs for critical operations</li>")
        sb.AppendLine("            </ul>")
        sb.AppendLine("        </div>")

        ' Output Management
        sb.AppendLine("        <div class='feature-section'>")
        sb.AppendLine("            <h2>6. Output Management</h2>")
        sb.AppendLine("            <ul>")
        sb.AppendLine("                <li>Timestamped output folders for organization</li>")
        sb.AppendLine("                <li>Automatic folder creation if needed</li>")
        sb.AppendLine("                <li>Preservation of original document formatting</li>")
        sb.AppendLine("                <li>Direct access to output folders</li>")
        sb.AppendLine("                <li>Default and custom output path options</li>")
        sb.AppendLine("            </ul>")
        sb.AppendLine("        </div>")

        ' Memory Management
        sb.AppendLine("        <div class='feature-section'>")
        sb.AppendLine("            <h2>7. Memory Management</h2>")
        sb.AppendLine("            <ul>")
        sb.AppendLine("                <li>Efficient COM object lifecycle management</li>")
        sb.AppendLine("                <li>Automatic resource cleanup</li>")
        sb.AppendLine("                <li>Proper disposal of Excel and Word processes</li>")
        sb.AppendLine("                <li>Temporary file management</li>")
        sb.AppendLine("                <li>Memory optimization during processing</li>")
        sb.AppendLine("            </ul>")
        sb.AppendLine("        </div>")

        ' New Features
        sb.AppendLine("        <div class='feature-section'>")
        sb.AppendLine("            <h2>8. Latest Enhancements</h2>")
        sb.AppendLine("            <ul>")
        sb.AppendLine("                <li>Enhanced error recovery system</li>")
        sb.AppendLine("                <li>Improved progress tracking with detailed status updates</li>")
        sb.AppendLine("                <li>Quick access to output folders</li>")
        sb.AppendLine("                <li>Confirmation dialogs for critical operations</li>")
        sb.AppendLine("                <li>Advanced data type preservation</li>")
        sb.AppendLine("                <li>Comprehensive help documentation</li>")
        sb.AppendLine("            </ul>")
        sb.AppendLine("        </div>")

        ' Footer
        sb.AppendLine("        <div class='footer'>")
        sb.AppendLine("            <p>Generated on: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & "</p>")
        sb.AppendLine("            <p>Auto Fill Documents Application - Version 1.1</p>")
        sb.AppendLine("            <p>? " & DateTime.Now.Year.ToString() & " All Rights Reserved</p>")
        sb.AppendLine("        </div>")
        sb.AppendLine("    </div>")
        sb.AppendLine("</body>")
        sb.AppendLine("</html>")

        Return sb.ToString()
    End Function

    Private Sub ShowControlsHelp()
        Try
            Dim tempPath As String = Environment.GetEnvironmentVariable("TEMP")
            Dim htmlFile As String = IO.Path.Combine(tempPath, "AutoFillDocsControls.html")

            ' Generate HTML content
            Dim htmlContent As String = GenerateControlsHelpHtml()
            ' Save to temp file
            IO.File.WriteAllText(htmlFile, htmlContent, System.Text.Encoding.UTF8)

            ' Open in browser
            Reusable_Class_Functions.OpenFileInBrowser(htmlFile)

        Catch ex As Exception
            MessageBox.Show($"Error showing controls help: {ex.Message}", "Help Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function GenerateControlsHelpHtml() As String
        Dim sb As New Text.StringBuilder()

        sb.AppendLine("<!DOCTYPE html>")
        sb.AppendLine("<html lang='en'>")
        sb.AppendLine("<head>")
        sb.AppendLine("    <meta charset='UTF-8'>")
        sb.AppendLine("    <meta name='viewport' content='width=device-width, initial-scale=1.0'>")
        sb.AppendLine("    <title>Auto Fill Documents - Help Guide</title>")
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
        sb.AppendLine("        .control-group {")
        sb.AppendLine("            margin-bottom: 20px;")
        sb.AppendLine("            padding: 15px;")
        sb.AppendLine("            border: 1px solid #e0e0e0;")
        sb.AppendLine("            border-radius: 5px;")
        sb.AppendLine("        }")
        sb.AppendLine("        .control-name {")
        sb.AppendLine("            font-weight: bold;")
        sb.AppendLine("            color: #2c3e50;")
        sb.AppendLine("        }")
        sb.AppendLine("        .control-description {")
        sb.AppendLine("            margin-left: 20px;")
        sb.AppendLine("        }")
        sb.AppendLine("        .footer {")
        sb.AppendLine("            text-align: center;")
        sb.AppendLine("            margin-top: 30px;")
        sb.AppendLine("            color: #7f8c8d;")
        sb.AppendLine("            font-size: 0.9em;")
        sb.AppendLine("        }")
        sb.AppendLine("        .logic-flow {")
        sb.AppendLine("            margin: 20px 0;")
        sb.AppendLine("            padding: 15px;")
        sb.AppendLine("            background-color: #f8f9fa;")
        sb.AppendLine("            border-radius: 5px;")
        sb.AppendLine("        }")
        sb.AppendLine("        .logic-step {")
        sb.AppendLine("            margin: 10px 0;")
        sb.AppendLine("            padding: 10px;")
        sb.AppendLine("            background-color: #fff;")
        sb.AppendLine("            border-left: 4px solid #3498db;")
        sb.AppendLine("        }")
        sb.AppendLine("        .logic-branch {")
        sb.AppendLine("            margin-left: 20px;")
        sb.AppendLine("            border-left: 2px dashed #3498db;")
        sb.AppendLine("            padding-left: 15px;")
        sb.AppendLine("        }")
        sb.AppendLine("    </style>")
        sb.AppendLine("</head>")
        sb.AppendLine("<body>")
        sb.AppendLine("    <div class='container'>")
        sb.AppendLine("        <h1>Auto Fill Documents - Help Guide</h1>")

        ' Software Logic Flow Section
        sb.AppendLine("        <h2>Software Logic Flow</h2>")
        sb.AppendLine("        <div class='logic-flow'>")
        sb.AppendLine("            <h3>1. Initialization</h3>")
        sb.AppendLine("            <div class='logic-step'>")
        sb.AppendLine("                ? Application starts and initializes the main form")
        sb.AppendLine("                ? ComboBoxes are set to DropDownList style")
        sb.AppendLine("                ? Default output path is set")
        sb.AppendLine("                ? Control states are initialized based on checkboxes")
        sb.AppendLine("            </div>")

        sb.AppendLine("            <h3>2. Document Selection Process</h3>")
        sb.AppendLine("            <div class='logic-step'>")
        sb.AppendLine("                <strong>Word Documents:</strong>")
        sb.AppendLine("                <div class='logic-branch'>")
        sb.AppendLine("                    ? User selects Word document(s)")
        sb.AppendLine("                    ? System scans for merge fields (<<FieldName>>)")
        sb.AppendLine("                    ? Fields are populated in respective ComboBoxes")
        sb.AppendLine("                    ? Validation ensures no duplicate selections")
        sb.AppendLine("                </div>")
        sb.AppendLine("                <strong>Excel Files:</strong>")
        sb.AppendLine("                <div class='logic-branch'>")
        sb.AppendLine("                    ? User selects Excel file(s)")
        sb.AppendLine("                    ? System reads headers from first row")
        sb.AppendLine("                    ? Headers are populated in respective ComboBoxes")
        sb.AppendLine("                    ? Validation prevents duplicate file selection")
        sb.AppendLine("                </div>")
        sb.AppendLine("            </div>")

        sb.AppendLine("            <h3>3. Field Mapping</h3>")
        sb.AppendLine("            <div class='logic-step'>")
        sb.AppendLine("                ? Auto-populate button matches fields automatically")
        sb.AppendLine("                ? Manual field matching dialog appears if needed")
        sb.AppendLine("                ? System validates field selections")
        sb.AppendLine("                ? Join fields are selected for data verification")
        sb.AppendLine("            </div>")

        sb.AppendLine("            <h3>4. Data Processing Options</h3>")
        sb.AppendLine("            <div class='logic-step'>")
        sb.AppendLine("                <strong>Document Generation:</strong>")
        sb.AppendLine("                <div class='logic-branch'>")
        sb.AppendLine("                    ? Separate documents per row")
        sb.AppendLine("                    ? Single document with page breaks")
        sb.AppendLine("                </div>")
        sb.AppendLine("                <strong>Data Verification:</strong>")
        sb.AppendLine("                <div class='logic-branch'>")
        sb.AppendLine("                    ? Join data from multiple Excel files")
        sb.AppendLine("                    ? Preview joined data")
        sb.AppendLine("                    ? Validate data relationships")
        sb.AppendLine("                </div>")
        sb.AppendLine("            </div>")

        sb.AppendLine("            <h3>5. Processing and Output</h3>")
        sb.AppendLine("            <div class='logic-step'>")
        sb.AppendLine("                ? Creates timestamped output folder")
        sb.AppendLine("                ? Processes each Excel row")
        sb.AppendLine("                ? Replaces merge fields with data")
        sb.AppendLine("                ? Maintains formatting and data types")
        sb.AppendLine("                ? Provides progress updates")
        sb.AppendLine("                ? Generates output documents")
        sb.AppendLine("            </div>")

        sb.AppendLine("            <h3>6. Error Handling and Validation</h3>")
        sb.AppendLine("            <div class='logic-step'>")
        sb.AppendLine("                ? Validates file selections")
        sb.AppendLine("                ? Ensures required fields are selected")
        sb.AppendLine("                ? Handles COM object lifecycle")
        sb.AppendLine("                ? Provides user feedback for errors")
        sb.AppendLine("                ? Maintains data integrity")
        sb.AppendLine("            </div>")
        sb.AppendLine("        </div>")

        ' Original Controls Sections
        sb.AppendLine("        <h2>Original Controls</h2>")
        sb.AppendLine("        <div class='control-group'>")
        sb.AppendLine("            <p class='control-name'>Word Document Selection Buttons (Top Section)</p>")
        sb.AppendLine("            <p class='control-description'>Located in the upper section, these buttons open file dialogs to select Word documents (.doc, .docx) for mail merge templates.</p>")
        sb.AppendLine("        </div>")
        sb.AppendLine("        <div class='control-group'>")
        sb.AppendLine("            <p class='control-name'>Excel File Selection Buttons (Middle Section)</p>")
        sb.AppendLine("            <p class='control-description'>Located in the middle section, these buttons open file dialogs to select Excel files (.xls, .xlsx, .xlsm) containing data for the mail merge.</p>")
        sb.AppendLine("        </div>")
        sb.AppendLine("        <div class='control-group'>")
        sb.AppendLine("            <p class='control-name'>Process Button (Bottom Right)</p>")
        sb.AppendLine("            <p class='control-description'>Located at the bottom right, initiates the mail merge process using the selected documents and settings.</p>")
        sb.AppendLine("        </div>")
        sb.AppendLine("        <div class='control-group'>")
        sb.AppendLine("            <p class='control-name'>Close Button (Bottom)</p>")
        sb.AppendLine("            <p class='control-description'>Located at the bottom, closes the application after cleaning up resources.</p>")
        sb.AppendLine("        </div>")
        sb.AppendLine("        <div class='control-group'>")
        sb.AppendLine("            <p class='control-name'>Auto-Populate Button (Center)</p>")
        sb.AppendLine("            <p class='control-description'>Located in the center section, automatically populates the field mapping dropdowns with field names from selected documents.</p>")
        sb.AppendLine("        </div>")
        sb.AppendLine("        <div class='control-group'>")
        sb.AppendLine("            <p class='control-name'>Output Folder Selection Button (Bottom Section)</p>")
        sb.AppendLine("            <p class='control-description'>Located in the bottom section, opens a folder dialog to select the destination for generated documents.</p>")
        sb.AppendLine("        </div>")

        ' ComboBoxes Section
        sb.AppendLine("        <h2>Field Mapping Dropdowns</h2>")
        sb.AppendLine("        <div class='control-group'>")
        sb.AppendLine("            <p class='control-name'>Word Document Field Dropdowns (Upper Section)</p>")
        sb.AppendLine("            <p class='control-description'>Located in the upper section, displays field names found in the selected Word documents.</p>")
        sb.AppendLine("        </div>")
        sb.AppendLine("        <div class='control-group'>")
        sb.AppendLine("            <p class='control-name'>Primary Excel Field Dropdowns (Middle Section)</p>")
        sb.AppendLine("            <p class='control-description'>Located in the middle section, displays headers from selected Excel files for primary field mapping.</p>")
        sb.AppendLine("        </div>")
        sb.AppendLine("        <div class='control-group'>")
        sb.AppendLine("            <p class='control-name'>Additional Excel Field Dropdowns (Lower Section)</p>")
        sb.AppendLine("            <p class='control-description'>Located in the lower section, used for additional field mapping when multiple Excel files are active.</p>")
        sb.AppendLine("        </div>")

        ' CheckBoxes Section
        sb.AppendLine("        <h2>Document Control Checkboxes</h2>")
        sb.AppendLine("        <div class='control-group'>")
        sb.AppendLine("            <p class='control-name'>Word Document Ignore Checkboxes (Upper Section)</p>")
        sb.AppendLine("            <p class='control-description'>Located next to Word document selection controls, when checked, excludes the corresponding Word document from processing.</p>")
        sb.AppendLine("        </div>")
        sb.AppendLine("        <div class='control-group'>")
        sb.AppendLine("            <p class='control-name'>Excel File Ignore Checkboxes (Middle Section)</p>")
        sb.AppendLine("            <p class='control-description'>Located next to Excel file selection controls, when checked, excludes the corresponding Excel file from processing.</p>")
        sb.AppendLine("        </div>")
        sb.AppendLine("        <div class='control-group'>")
        sb.AppendLine("            <p class='control-name'>Default Output Path Checkbox (Bottom Section)</p>")
        sb.AppendLine("            <p class='control-description'>Located in the output section, when checked, uses the default output path (D:\AUTOFILL_DOCS or Documents\AUTOFILL_DOCS).</p>")
        sb.AppendLine("        </div>")

        ' RadioButtons Section
        sb.AppendLine("        <h2>Output Format Options</h2>")
        sb.AppendLine("        <div class='control-group'>")
        sb.AppendLine("            <p class='control-name'>Document Generation Options (Right Section)</p>")
        sb.AppendLine("            <p class='control-description'>Choose between generating separate documents for each data row or a single combined document with page breaks between entries.</p>")
        sb.AppendLine("        </div>")

        ' TextBoxes Section
        sb.AppendLine("        <h2>File Path Display Fields</h2>")
        sb.AppendLine("        <div class='control-group'>")
        sb.AppendLine("            <p class='control-name'>Word Document Paths (Upper Section)</p>")
        sb.AppendLine("            <p class='control-description'>Located in the Word documents section, displays the paths of selected Word template files.</p>")
        sb.AppendLine("        </div>")
        sb.AppendLine("        <div class='control-group'>")
        sb.AppendLine("            <p class='control-name'>Excel File Paths (Middle Section)</p>")
        sb.AppendLine("            <p class='control-description'>Located in the Excel files section, displays the paths of selected data source files.</p>")
        sb.AppendLine("        </div>")
        sb.AppendLine("        <div class='control-group'>")
        sb.AppendLine("            <p class='control-name'>Output Folder Path (Bottom Section)</p>")
        sb.AppendLine("            <p class='control-description'>Located in the output section, displays the destination folder path for generated documents.</p>")
        sb.AppendLine("        </div>")

        ' Progress Indicators Section
        sb.AppendLine("        <h2>Progress Indicators</h2>")
        sb.AppendLine("        <div class='control-group'>")
        sb.AppendLine("            <p class='control-name'>Progress Bar (Bottom Section)</p>")
        sb.AppendLine("            <p class='control-description'>Located at the bottom of the window, shows the progress of current operations (file reading, processing, etc.).</p>")
        sb.AppendLine("        </div>")
        sb.AppendLine("        <div class='control-group'>")
        sb.AppendLine("            <p class='control-name'>Status Label (Bottom Section)</p>")
        sb.AppendLine("            <p class='control-description'>Located next to the progress bar, displays current operation status and progress messages.</p>")
        sb.AppendLine("        </div>")

        ' Footer
        sb.AppendLine("        <div class='footer'>")
        sb.AppendLine("            <p>Generated on: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & "</p>")
        sb.AppendLine("        </div>")
        sb.AppendLine("    </div>")
        sb.AppendLine("</body>")
        sb.AppendLine("</html>")

        Return sb.ToString()
    End Function

    Private Sub HelpToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HelpToolStripMenuItem.Click
        ShowControlsHelp()
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        ShowAboutInfo()
    End Sub

    Private Sub DataVerificationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DataVerificationToolStripMenuItem.Click
        ' Track active Excel files and their selected join fields
        Dim activeFiles As New List(Of Tuple(Of String, ComboBox))()

        ' For Excel file 1 (TextBox3), use ComboBox3 for primary field and ComboBox6 for join field
        If Not CheckBox3.Checked AndAlso Not String.IsNullOrEmpty(TextBox3.Text) Then
            activeFiles.Add(Tuple.Create(TextBox3.Text, ComboBox3))
        End If

        ' For Excel file 2 (TextBox4), use ComboBox4 for primary field and ComboBox7 for join field
        If Not CheckBox4.Checked AndAlso Not String.IsNullOrEmpty(TextBox4.Text) Then
            activeFiles.Add(Tuple.Create(TextBox4.Text, ComboBox4))
        End If

        ' For Excel file 3 (TextBox5), use ComboBox5 for primary field and ComboBox8 for join field
        If Not CheckBox5.Checked AndAlso Not String.IsNullOrEmpty(TextBox5.Text) Then
            activeFiles.Add(Tuple.Create(TextBox5.Text, ComboBox5))
        End If

        ' Verify we have at least two files for joining
        If activeFiles.Count < 2 Then
            MessageBox.Show("Please select at least two Excel files to verify data joins.",
                          "Insufficient Files", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If


    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        ' Show confirmation dialog
        If MessageBox.Show("Do you want to process the selected documents?", "Confirm Processing",
                          MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            Return
        End If

        Dim CountExcelFileUploaded As Integer
        Dim ExcelPath As String
        Try
            ' Validate at least one document is selected and not ignored
            If (CheckBox1.Checked OrElse String.IsNullOrEmpty(TextBox1.Text)) AndAlso
               (CheckBox2.Checked OrElse String.IsNullOrEmpty(TextBox2.Text)) Then
                MessageBox.Show("Please select at least one Word document to process.", "No Documents Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' Validate at least one Excel file is selected and not ignored
            If (CheckBox3.Checked OrElse String.IsNullOrEmpty(TextBox3.Text)) AndAlso
               (CheckBox4.Checked OrElse String.IsNullOrEmpty(TextBox4.Text)) AndAlso
               (CheckBox5.Checked OrElse String.IsNullOrEmpty(TextBox5.Text)) Then
                MessageBox.Show("Please select at least one Excel file to process.", "No Data Source", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' Validate output path
            If String.IsNullOrEmpty(TextBox6.Text) Then
                MessageBox.Show("Please select an output folder.", "No Output Folder", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' Create output directory if it doesn't exist
            If Not IO.Directory.Exists(TextBox6.Text) Then
                IO.Directory.CreateDirectory(TextBox6.Text)
            End If
            CountExcelFileUploaded = 0
            ExcelPath = ""
            ' Process each active Excel file
            If Not CheckBox3.Checked AndAlso Not String.IsNullOrEmpty(TextBox3.Text) Then
                CountExcelFileUploaded += 1
                ExcelPath = TextBox3.Text
            End If
            If Not CheckBox4.Checked AndAlso Not String.IsNullOrEmpty(TextBox4.Text) Then
                CountExcelFileUploaded += 1
                ExcelPath = TextBox4.Text
            End If
            If Not CheckBox5.Checked AndAlso Not String.IsNullOrEmpty(TextBox5.Text) Then
                CountExcelFileUploaded += 1
                ExcelPath = TextBox5.Text
            End If

            If CountExcelFileUploaded = 1 Then
                ProcessMailMerge(ExcelPath)

            ElseIf CountExcelFileUploaded > 1 Then
                If (String.IsNullOrEmpty(Txt_Join.Text)) Then
                    MessageBox.Show("Please validate joined Excel files before processing (Press Validate Button).", "No Data Source", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Btn_Join.Focus()
                    Return
                End If
                ProcessMailMerge(Txt_Join.Text)
            Else
                MessageBox.Show("Please upload proper Excel files and document files! Process failed.", "Process Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

        Catch ex As Exception
            MessageBox.Show("Error processing documents: " & ex.Message, "Processing Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Btn_Join_Click(sender As Object, e As EventArgs) Handles Btn_Join.Click
        ' Show confirmation dialog
        If MessageBox.Show("Do you want to join the selected Excel files?", "Confirm Join Operation",
                          MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            Return
        End If

        Try
            UpdateProgress(0, "Initializing join operation...")

            ' Validate that at least two Excel files are selected and not ignored
            Dim activeFiles As New List(Of Tuple(Of String, String, ComboBox))() ' (FilePath, Prefix, ComboBox)

            ' Add active files with their prefixes
            If Not CheckBox3.Checked AndAlso Not String.IsNullOrEmpty(TextBox3.Text) Then
                activeFiles.Add(Tuple.Create(TextBox3.Text, "xls1", ComboBox3))
            End If
            If Not CheckBox4.Checked AndAlso Not String.IsNullOrEmpty(TextBox4.Text) Then
                activeFiles.Add(Tuple.Create(TextBox4.Text, "xls2", ComboBox4))
            End If
            If Not CheckBox5.Checked AndAlso Not String.IsNullOrEmpty(TextBox5.Text) Then
                activeFiles.Add(Tuple.Create(TextBox5.Text, "xls3", ComboBox5))
            End If

            ' Verify we have at least two files
            If activeFiles.Count < 2 Then
                MessageBox.Show("Please select at least two Excel files to join.",
                              "Insufficient Files", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' Verify field selections
            For Each file In activeFiles
                If file.Item3.SelectedIndex = -1 Then
                    MessageBox.Show($"Please select a join field for {IO.Path.GetFileName(file.Item1)}",
                                  "Field Selection Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If
            Next

            UpdateProgress(10, "Creating output directory...")

            ' Create output directory in temp folder
            Dim tempPath As String = Environment.GetEnvironmentVariable("TEMP")
            Dim joinPath As String = IO.Path.Combine(tempPath, "JOIN_XLS")
            If Not IO.Directory.Exists(joinPath) Then
                IO.Directory.CreateDirectory(joinPath)
            End If

            ' Determine base file based on radio button selection
            Dim baseFile As Tuple(Of String, String, ComboBox) = Nothing
            If RadioButton3.Checked Then
                baseFile = activeFiles.Find(Function(f) f.Item1 = TextBox3.Text)
            ElseIf RadioButton4.Checked Then
                baseFile = activeFiles.Find(Function(f) f.Item1 = TextBox4.Text)
            ElseIf RadioButton5.Checked Then
                baseFile = activeFiles.Find(Function(f) f.Item1 = TextBox5.Text)
            End If

            If baseFile Is Nothing Then
                MessageBox.Show("Please select a base file using the radio buttons.",
                              "Base File Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            UpdateProgress(20, "Opening Excel files...")

            ' Create Excel objects
            Dim excel As Object = Nothing
            Dim baseWorkbook As Object = Nothing
            Dim resultWorkbook As Object = Nothing

            Try
                excel = CreateObject("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False

                UpdateProgress(30, "Reading base file...")

                ' Open base workbook and get its data
                baseWorkbook = excel.Workbooks.Open(baseFile.Item1)
                Dim baseSheet = baseWorkbook.Worksheets(1)
                Dim baseRange = baseSheet.UsedRange
                Dim baseHeaders = GetExcelHeaders(baseFile.Item1)
                Dim baseJoinField = baseFile.Item3.SelectedItem.ToString()
                Dim baseJoinIndex = baseHeaders.IndexOf(baseJoinField) + 1

                UpdateProgress(40, "Creating result workbook...")

                ' Create new workbook for results
                resultWorkbook = excel.Workbooks.Add()
                Dim resultSheet = resultWorkbook.Worksheets(1)

                ' Build header mapping and write headers
                Dim headerMapping As New Dictionary(Of String, String)() ' Original -> Prefixed
                Dim colIndex As Integer = 1

                ' Add base file headers first with prefixes
                For Each header In baseHeaders
                    Dim newHeader As String = $"{baseFile.Item2}.{header}"
                    headerMapping.Add(header, newHeader)
                    resultSheet.Cells(1, colIndex) = newHeader
                    ' Copy the number format from base file to maintain data type
                    resultSheet.Cells(1, colIndex).NumberFormat = baseSheet.Cells(1, baseHeaders.IndexOf(header) + 1).NumberFormat
                    colIndex += 1
                Next

                UpdateProgress(50, "Processing additional files...")

                ' Process each additional file
                Dim otherFiles = activeFiles.Where(Function(f) f.Item1 <> baseFile.Item1).ToList()
                Dim progressPerFile = 40 / otherFiles.Count
                Dim currentProgress = 50
                Dim totalMatchingRecords As Integer = 0

                For Each otherFile In otherFiles
                    UpdateProgress(currentProgress, $"Processing {IO.Path.GetFileName(otherFile.Item1)}...")

                    Dim otherWorkbook As Object = Nothing
                    Try
                        ' Open other workbook
                        otherWorkbook = excel.Workbooks.Open(otherFile.Item1)
                        Dim otherSheet = otherWorkbook.Worksheets(1)
                        Dim otherRange = otherSheet.UsedRange
                        Dim otherHeaders = GetExcelHeaders(otherFile.Item1)
                        Dim otherJoinField = otherFile.Item3.SelectedItem.ToString()
                        Dim otherJoinIndex = otherHeaders.IndexOf(otherJoinField) + 1

                        ' Add other file headers with prefixes
                        For Each header In otherHeaders
                            Dim newHeader As String = $"{otherFile.Item2}.{header}"
                            headerMapping.Add($"{otherFile.Item2}_{header}", newHeader)
                            resultSheet.Cells(1, colIndex) = newHeader
                            ' Copy the number format from source file to maintain data type
                            resultSheet.Cells(1, colIndex).NumberFormat = otherSheet.Cells(1, otherHeaders.IndexOf(header) + 1).NumberFormat
                            colIndex += 1
                        Next

                        ' Copy matching data
                        Dim baseRowCount = baseRange.Rows.Count
                        Dim matchesFound As Integer = 0

                        For baseRow = 2 To baseRowCount
                            Dim baseValue = baseSheet.Cells(baseRow, baseJoinIndex).Value

                            ' Find matching row in other file
                            Dim found = False
                            For otherRow = 2 To otherRange.Rows.Count
                                If otherSheet.Cells(otherRow, otherJoinIndex).Value = baseValue Then
                                    found = True
                                    matchesFound += 1
                                    ' Copy data from other file
                                    Dim otherColIndex = colIndex - otherHeaders.Count
                                    For Each header In otherHeaders
                                        Dim sourceCol = otherHeaders.IndexOf(header) + 1
                                        Dim value = otherSheet.Cells(otherRow, sourceCol).Value
                                        resultSheet.Cells(baseRow, otherColIndex) = value
                                        ' Copy the number format to maintain data type
                                        resultSheet.Cells(baseRow, otherColIndex).NumberFormat = otherSheet.Cells(otherRow, sourceCol).NumberFormat
                                        otherColIndex += 1
                                    Next
                                    Exit For
                                End If
                            Next

                            If Not found Then
                                ' Fill empty cells for non-matching records
                                Dim otherColIndex = colIndex - otherHeaders.Count
                                For i = 1 To otherHeaders.Count
                                    resultSheet.Cells(baseRow, otherColIndex + i - 1) = ""
                                Next
                            End If
                        Next

                        If matchesFound = 0 Then
                            MessageBox.Show($"No matching records found between base file and {IO.Path.GetFileName(otherFile.Item1)} using join field '{otherJoinField}'",
                                          "No Matches Found", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Else
                            totalMatchingRecords += matchesFound
                        End If

                    Finally
                        If otherWorkbook IsNot Nothing Then
                            otherWorkbook.Close(False)
                            ReleaseComObject(otherWorkbook)
                        End If
                    End Try

                    currentProgress += progressPerFile
                Next

                UpdateProgress(90, "Saving joined data...")

                ' Copy base data with original data types
                For row = 2 To baseRange.Rows.Count
                    For col = 1 To baseHeaders.Count
                        resultSheet.Cells(row, col) = baseSheet.Cells(row, col).Value
                        ' Copy the number format to maintain data type
                        resultSheet.Cells(row, col).NumberFormat = baseSheet.Cells(row, col).NumberFormat
                    Next
                Next

                ' Save result workbook
                Dim timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss")
                Dim resultPath = IO.Path.Combine(joinPath, $"JoinedData_{timestamp}.xlsx")
                resultWorkbook.SaveAs(resultPath)

                ' Update UI
                Txt_Join.Text = resultPath

                ' Update ComboBox with all headers
                Cbx_Join.Items.Clear()
                Cbx_Join.Items.AddRange(headerMapping.Values.ToArray())
                If Cbx_Join.Items.Count > 0 Then
                    Cbx_Join.SelectedIndex = 0
                End If

                UpdateProgress(100, "Join operation completed successfully!")
                MessageBox.Show($"Excel files have been joined successfully with {totalMatchingRecords} total matching records!",
                              "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Finally
                ' Cleanup
                If resultWorkbook IsNot Nothing Then
                    resultWorkbook.Close(True)
                    ReleaseComObject(resultWorkbook)
                End If
                If baseWorkbook IsNot Nothing Then
                    baseWorkbook.Close(False)
                    ReleaseComObject(baseWorkbook)
                End If
                If excel IsNot Nothing Then
                    excel.Quit()
                    ReleaseComObject(excel)
                End If
            End Try
            MessageBox.Show("Validation successful. Excel and Word files are ready to process.", "Validation Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            UpdateProgress(100, "Error during join operation!")
            MessageBox.Show($"Error joining Excel files: {ex.Message}", "Join Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' Reset progress bar after a delay
            System.Threading.Thread.Sleep(1000)
            UpdateProgress(0, "")
        End Try
    End Sub

    ' Modify Button8_Click (Close) to use enhanced error handling
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If MessageBox.Show("Do you want to close the application? This will clean up all temporary files and release resources.",
                          "Confirm Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            Return
        End If

        Try
            ' Initial cleanup
            UpdateProgress(0, "Initializing cleanup process...")

            ' Memory cleanup
            UpdateProgress(10, "Cleaning up memory resources...")
            GC.Collect()
            GC.WaitForPendingFinalizers()

            ' Clean up COM processes
            UpdateProgress(20, "Closing Office applications...")
            SafeKillProcess("EXCEL")
            SafeKillProcess("WINWORD")

            ' Clean up temp files
            UpdateProgress(40, "Cleaning temporary files...")
            Dim tempPath As String = Environment.GetEnvironmentVariable("TEMP")
            If Not String.IsNullOrEmpty(tempPath) Then
                ' Clean app-specific files
                SafeFileOperation(Sub()
                                      Dim joinPath = IO.Path.Combine(tempPath, "JOIN_XLS")
                                      If IO.Directory.Exists(joinPath) Then IO.Directory.Delete(joinPath, True)

                                      Dim helpFile = IO.Path.Combine(tempPath, "AutoFillDocsControls.html")
                                      Dim aboutFile = IO.Path.Combine(tempPath, "AutoFillDocsSynopsis.html")

                                      If IO.File.Exists(helpFile) Then IO.File.Delete(helpFile)
                                      If IO.File.Exists(aboutFile) Then IO.File.Delete(aboutFile)
                                  End Sub, "temporary file cleanup")
            End If

            ' Final cleanup
            UpdateProgress(80, "Performing final cleanup...")
            GC.Collect()
            GC.WaitForPendingFinalizers()

            UpdateProgress(100, "Cleanup complete!")
            System.Threading.Thread.Sleep(500)

            ' Exit application
            Application.Exit()

        Catch ex As Exception
            HandleError(ex, "application shutdown")
            ' Force exit if cleanup fails
            Application.Exit()
        End Try
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        ' Show confirmation dialog
        If MessageBox.Show("Do you want to reset all settings to default values? This will clear all selections and temporary files.",
                          "Confirm Reset", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.No Then
            Return
        End If

        Try
            ' Initial cleanup
            UpdateProgress(0, "Initializing reset process...")
            System.Threading.Thread.Sleep(500)

            ' Memory cleanup first
            UpdateProgress(10, "Cleaning up memory resources...")
            GC.Collect()
            GC.WaitForPendingFinalizers()
            System.Threading.Thread.Sleep(500)

            ' Clean up any open COM objects
            UpdateProgress(15, "Releasing COM objects...")
            Try
                For Each proc As Process In Process.GetProcessesByName("EXCEL")
                    proc.Kill()
                Next
                For Each proc As Process In Process.GetProcessesByName("WINWORD")
                    proc.Kill()
                Next
            Catch ex As Exception
                ' Silently continue if we can't kill processes
            End Try
            System.Threading.Thread.Sleep(500)

            ' Reset Word document section
            UpdateProgress(20, "Resetting Word document selections...")
            TextBox1.Clear()
            TextBox2.Clear()
            ComboBox1.Items.Clear()
            ComboBox2.Items.Clear()
            CheckBox1.Checked = False
            CheckBox2.Checked = False
            System.Threading.Thread.Sleep(500)

            ' Reset Excel file section
            UpdateProgress(35, "Resetting Excel file selections...")
            TextBox3.Clear()
            TextBox4.Clear()
            TextBox5.Clear()
            ComboBox3.Items.Clear()
            ComboBox4.Items.Clear()
            ComboBox5.Items.Clear()
            CheckBox3.Checked = False
            CheckBox4.Checked = False
            CheckBox5.Checked = False
            System.Threading.Thread.Sleep(500)

            ' Reset join section
            UpdateProgress(50, "Resetting join settings...")
            Txt_Join.Clear()
            Cbx_Join.Items.Clear()
            RadioButton3.Checked = True ' Set first radio button as default
            RadioButton4.Checked = False
            RadioButton5.Checked = False
            System.Threading.Thread.Sleep(500)

            ' Clean temporary files
            UpdateProgress(65, "Cleaning temporary files...")
            Try
                ' Get temp directory path
                Dim tempPath As String = Environment.GetEnvironmentVariable("TEMP")

                ' Clean up join files
                Dim joinPath As String = IO.Path.Combine(tempPath, "JOIN_XLS")
                If IO.Directory.Exists(joinPath) Then
                    IO.Directory.Delete(joinPath, True)
                End If

                ' Clean up help and about files
                Dim helpFile As String = IO.Path.Combine(tempPath, "AutoFillDocsControls.html")
                Dim aboutFile As String = IO.Path.Combine(tempPath, "AutoFillDocsSynopsis.html")

                If IO.File.Exists(helpFile) Then IO.File.Delete(helpFile)
                If IO.File.Exists(aboutFile) Then IO.File.Delete(aboutFile)

            Catch ex As Exception
                ' Silently continue if files are locked or inaccessible
            End Try
            System.Threading.Thread.Sleep(500)

            ' Reset output options
            UpdateProgress(80, "Resetting output options...")
            RadioButton1.Checked = True ' Set "Separate Documents" as default
            RadioButton2.Checked = False
            CheckBox6.Checked = True ' Set default output path
            TextBox6.Text = Reusable_Class_Functions.GetDefaultOutputPath()
            System.Threading.Thread.Sleep(500)

            ' Reset GroupBox states
            UpdateProgress(90, "Resetting control states...")
            InitializeControlStates()
            UpdateGroupBox4State()
            System.Threading.Thread.Sleep(500)

            ' Final memory cleanup
            UpdateProgress(95, "Performing final memory cleanup...")
            GC.Collect()
            GC.WaitForPendingFinalizers()
            System.Threading.Thread.Sleep(500)

            UpdateProgress(100, "Reset complete!")

            MessageBox.Show("All settings have been reset to default values and memory has been cleared.", "Reset Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            UpdateProgress(100, "Error during reset!")
            MessageBox.Show($"Error resetting application: {ex.Message}", "Reset Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' Reset progress bar after a delay
            System.Threading.Thread.Sleep(1000)
            UpdateProgress(0, "")
        End Try
    End Sub

    ' Add error recovery for file operations
    Private Function RecoverableFileOperation(Of T)(operation As Func(Of T), maxRetries As Integer) As T
        Dim attempt As Integer = 0
        Dim lastException As Exception = Nothing

        While attempt < maxRetries
            Try
                Return operation()
            Catch ex As IO.IOException
                lastException = ex
                attempt += 1
                If attempt < maxRetries Then
                    System.Threading.Thread.Sleep(500) ' Wait before retry
                    GC.Collect() ' Try to free up resources
                End If
            End Try
        End While

        ' If we get here, all attempts failed
        HandleError(lastException, "file operation after " & maxRetries & " retries")
        Return Nothing
    End Function

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Try
            ' Check if TextBox6 has a value
            If String.IsNullOrEmpty(TextBox6.Text.Trim()) Then
                MessageBox.Show("Please select an output folder first.", "No Folder Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' Check if the path exists
            If Not IO.Directory.Exists(TextBox6.Text) Then
                Dim result = MessageBox.Show("The specified folder does not exist. Would you like to create it?",
                                           "Folder Not Found",
                                           MessageBoxButtons.YesNo,
                                           MessageBoxIcon.Question)
                If result = DialogResult.Yes Then
                    Try
                        IO.Directory.CreateDirectory(TextBox6.Text)
                    Catch ex As Exception
                        MessageBox.Show($"Could not create folder: {ex.Message}", "Folder Creation Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return
                    End Try
                Else
                    Return
                End If
            End If

            ' Open the folder in Windows Explorer
            Process.Start("explorer.exe", TextBox6.Text)

        Catch ex As Exception
            HandleError(ex, "opening output folder")
        End Try
    End Sub

    Private Sub Excel1ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Excel1ToolStripMenuItem.Click
        Reusable_Class_Functions.VerifyAndDisplayExcelContent(TextBox3.Text)
    End Sub

    Private Sub Excel2ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Excel2ToolStripMenuItem.Click
        Reusable_Class_Functions.VerifyAndDisplayExcelContent(TextBox4.Text)
    End Sub

    Private Sub Excel3ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Excel3ToolStripMenuItem.Click
        Reusable_Class_Functions.VerifyAndDisplayExcelContent(TextBox5.Text)
    End Sub

    Private Sub JoinExcelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles JoinExcelToolStripMenuItem.Click
        Reusable_Class_Functions.VerifyAndDisplayExcelContent(Txt_Join.Text)
    End Sub
End Class
