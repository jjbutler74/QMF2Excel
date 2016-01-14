Imports System.Runtime.InteropServices
Imports System.IO

Public Class frmQMF2Excel
    ' Screen Values
    Dim strScrMode As String
    Dim strScrBatchFile As String
    Dim strScrMainframeHost As String
    Dim strScrMainframeUserId As String
    Dim strScrMainframePassword As String
    Dim strScrMainframeFile As String
    Dim strScrExcelFolder As String
    Dim strScrExcelFile As String
    Dim strScrAutoFormat As String
    Dim strScrSkipTotals As String ' not yet on form
    Dim strScrOverwrite As String
    Dim strScrRememberScreen As String

    ' Log File Vars
    Dim strLogFile As String ' Folder and File of Text Log File
    Dim objLogFileWriter As StreamWriter
    Dim blnUseLogFile As Boolean

    ' Array of Mainframe Files to Convert (from Screen or Batch)
    Dim intNumberOfMainframeFiles As Integer
    Dim MainframeFileArray(100, 3)

    ' Vars to use data from Mainframe File Array
    Dim intCurrentMainframeFile As Integer
    Dim strMainframeFile As String
    Dim strExcelFile As String
    Dim strExcelFolder As String
    Dim blnNumeric2Text As Boolean

    ' Flag for Succesful Batch File Laod
    Dim blnSuccesfulBatchLaod As Boolean

    ' Flag for Succesful Try/Catch
    Dim blnSuccessfulTry As Boolean

    ' Flag for Succesful Mainframe File open
    Dim blnSuccesfulMainframeFileOpen As Boolean

    ' Stores Mainframe File Lines
    Dim strLine As String

    ' Mainframe Line Number
    Dim intLineNumber As Integer

    ' Var to read from Mainframe file 
    Dim sr As StreamReader

    ' Number of rows used for headings 
    Dim intExcelNumberOfHeadingRows As Integer

    ' Number of rows used for column headings 
    Dim intExcelColumnNameRow As Integer

    ' Max number of char in column 1, used to fortmat coulmn width in Excel when headers are present
    Dim intMaxCol1Chars As Integer

    ' Excel Vars
    Dim xlsApplication
    Dim xlsWorkBook
    Dim xlsWorkSheet
    Dim xlRng As Excel.Range

    ' *** Vars for Main data move ***
    Dim RawHeadingArray(12) ' Start with 1 / only can hold 10 rows
    Dim intNumberHeadingRows As Integer
    Dim blnHeadingsFound As Boolean

    Dim RawColumnNameArray(10) ' Start with 1
    Dim intNumberRawColumnNameRows As Integer
    Dim strDiv As String

    Dim DivArray(100, 2) ' Up to 100 Columns / Start and End Position
    Dim intNumberColumns
    Dim flgDivStatus As String ' S - Looking for Start Pos / E - Looking for End Pos

    Dim ColumnNames(100)

    Dim ExcelArray(65535, 100)
    Dim intOutputRow

    Dim intCnt As Integer
    Dim intCnt2 As Integer

    Dim strPreviousLine As String
    Dim blnFootersFound As Boolean
    Dim strDataItem As String

    Dim strEndFound As String

    Dim blnExcelCopyComplete As Boolean

    Dim ColumnTotals(100) ' jjb-total
    Dim intFirstExcelDataRow As Integer ' jjb-total
    Dim blnAnyTotals As Boolean ' jjb-total
    Dim intTotalRow As Integer ' jjb-total

    Private Sub frmQMF2Excel_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Dim vrs As Version
        'vrs = New Version(Application.ProductVersion)
        'Me.Text = "QMF2Excel - Version " & vrs.Major & "." & vrs.Minor 
        Me.Text = "QMF2Excel - Version 1.2"
        cmbMode.Items.Add("Use Screen Values")
        cmbMode.Items.Add("Use Batch File")

        Call LoadRegSettings()
    End Sub

    Private Sub cmbMode_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbMode.SelectedIndexChanged
        If cmbMode.Text = "Use Batch File" Then
            txtBatchFile.Enabled = True
            txtBatchFile.BackColor = Color.White
            btnBatchFile.Enabled = True
            lbl1.ForeColor = SystemColors.ControlText
            txtMainframeFile.Enabled = False
            txtMainframeFile.BackColor = SystemColors.ControlLight
            lbl2.ForeColor = SystemColors.GrayText
            txtExcelFolder.Enabled = False
            txtExcelFolder.BackColor = SystemColors.ControlLight
            btnExcelFolder.Enabled = False
            lbl3.ForeColor = SystemColors.GrayText
            txtExcelFile.Enabled = False
            txtExcelFile.BackColor = SystemColors.ControlLight
            lbl4.ForeColor = SystemColors.GrayText
            chkNumeric.Enabled = False
        Else
            txtBatchFile.Enabled = False
            txtBatchFile.BackColor = SystemColors.ControlLight
            btnBatchFile.Enabled = False
            lbl1.ForeColor = SystemColors.GrayText
            txtMainframeFile.Enabled = True
            txtMainframeFile.BackColor = Color.White
            lbl2.ForeColor = SystemColors.ControlText
            txtExcelFolder.Enabled = True
            txtExcelFolder.BackColor = Color.White
            btnExcelFolder.Enabled = True
            lbl3.ForeColor = SystemColors.ControlText
            txtExcelFile.Enabled = True
            txtExcelFile.BackColor = Color.White
            lbl4.ForeColor = SystemColors.ControlText
            chkNumeric.Enabled = True
        End If
    End Sub

    Private Sub btnBatchFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBatchFile.Click
        Dim BatchFile As New System.Windows.Forms.OpenFileDialog

        ' Descriptive text displayed above the tree view control in the dialog box
        BatchFile.Title = "Select the Batch File"
        BatchFile.DefaultExt = "xls"
        BatchFile.Filter = "Microsoft Excel (*.xls)|*.xls"

        ' Sets the folder where the browsing starts from
        If Trim(txtBatchFile.Text) <> "" Then
            BatchFile.InitialDirectory = txtBatchFile.Text
        Else
            BatchFile.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        End If

        Dim dlgResult As DialogResult = BatchFile.ShowDialog()

        If dlgResult = Windows.Forms.DialogResult.OK Then
            txtBatchFile.Text = BatchFile.FileName
        End If
    End Sub

    Private Sub btnExcelFolder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExcelFolder.Click
        Dim ExcelFolder As New System.Windows.Forms.FolderBrowserDialog

        ' Descriptive text displayed above the tree view control in the dialog box
        ExcelFolder.Description = "Select the Excel Folder"

        ' Sets the folder where the browsing starts from
        If Trim(txtExcelFolder.Text) <> "" Then
            ExcelFolder.SelectedPath = txtExcelFolder.Text
        Else
            ExcelFolder.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        End If

        ' Do not show the button for new folder
        ExcelFolder.ShowNewFolderButton = True

        Dim dlgResult As DialogResult = ExcelFolder.ShowDialog()

        If dlgResult = Windows.Forms.DialogResult.OK Then
            txtExcelFolder.Text = ExcelFolder.SelectedPath
        End If
    End Sub

    Private Sub btnSubmit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        Call SubmitInt()
        Call MainProcess()
        Call SubmitComplete()
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        End
    End Sub

    Private Sub LoadScreenValues()
        strScrMode = cmbMode.Text
        strScrBatchFile = txtBatchFile.Text
        strScrMainframeHost = txtMainframeHost.Text
        strScrMainframeUserId = txtMainframeUserId.Text
        strScrMainframePassword = txtMainframePassword.Text
        strScrMainframeFile = txtMainframeFile.Text
        strScrExcelFolder = txtExcelFolder.Text
        strScrExcelFile = txtExcelFile.Text
        strScrAutoFormat = chkNumeric.Checked
        strScrSkipTotals = True ' Will make reg option later
        strScrOverwrite = chkOverwrite.Checked
        strScrRememberScreen = chkRememberScreen.Checked
    End Sub

    Private Sub LoadRegSettings()
        If GetSetting("QMF2Excel", "Screen Values", "Remember Screen Values") = "True" Or GetSetting("QMF2Excel", "Screen Values", "Remember Screen Values") = "" Then
            cmbMode.Text = GetSetting("QMF2Excel", "Screen Values", "Mode", "Use Screen Values")
            txtBatchFile.Text = GetSetting("QMF2Excel", "Screen Values", "Batch Folder and File", My.Application.Info.DirectoryPath & "\TestBatch.xls")
            txtMainframeHost.Text = GetSetting("QMF2Excel", "Screen Values", "Mainframe Host", "")
            txtMainframeUserId.Text = GetSetting("QMF2Excel", "Screen Values", "Mainframe User Id", "")
            txtMainframePassword.Text = GetSetting("QMF2Excel", "Screen Values", "Mainframe Password", "")
            txtMainframeFile.Text = GetSetting("QMF2Excel", "Screen Values", "Mainframe File", "")
            txtExcelFolder.Text = GetSetting("QMF2Excel", "Screen Values", "Excel Folder", My.Application.Info.DirectoryPath & "\")
            txtExcelFile.Text = GetSetting("QMF2Excel", "Screen Values", "Excel File", "TestData.xls")
            chkNumeric.Checked = GetSetting("QMF2Excel", "Screen Values", "Auto Format", "True")
            chkOverwrite.Checked = GetSetting("QMF2Excel", "Screen Values", "Overwrite Existing Files", "True")
            chkRememberScreen.Checked = True
        Else
            cmbMode.Text = "Use Screen Values"
        End If
    End Sub

    Private Sub SaveRegSettings()
        If strScrRememberScreen = "True" Then
            SaveSetting("QMF2Excel", "Screen Values", "Mode", strScrMode)
            SaveSetting("QMF2Excel", "Screen Values", "Batch Folder and File", strScrBatchFile)
            SaveSetting("QMF2Excel", "Screen Values", "Mainframe Host", strScrMainframeHost)
            SaveSetting("QMF2Excel", "Screen Values", "Mainframe User Id", strScrMainframeUserId)
            SaveSetting("QMF2Excel", "Screen Values", "Mainframe Password", strScrMainframePassword)
            SaveSetting("QMF2Excel", "Screen Values", "Mainframe File", strScrMainframeFile)
            SaveSetting("QMF2Excel", "Screen Values", "Excel Folder", strScrExcelFolder)
            SaveSetting("QMF2Excel", "Screen Values", "Excel File", strScrExcelFile)
            SaveSetting("QMF2Excel", "Screen Values", "Auto Format", strScrAutoFormat)
            SaveSetting("QMF2Excel", "Screen Values", "Overwrite Existing Files", strScrOverwrite)
            SaveSetting("QMF2Excel", "Screen Values", "Remember Screen Values", "True")
        Else
            SaveSetting("QMF2Excel", "Screen Values", "Mode", "")
            SaveSetting("QMF2Excel", "Screen Values", "Batch Folder and File", "")
            SaveSetting("QMF2Excel", "Screen Values", "Mainframe File", "")
            SaveSetting("QMF2Excel", "Screen Values", "Excel Folder", "")
            SaveSetting("QMF2Excel", "Screen Values", "Excel File", "")
            SaveSetting("QMF2Excel", "Screen Values", "Auto Format", "")
            SaveSetting("QMF2Excel", "Screen Values", "Overwrite Existing Files", "")
            SaveSetting("QMF2Excel", "Screen Values", "Remember Screen Values", "False")
        End If
    End Sub

    Private Sub SubmitInt()
        ' Init
        btnSubmit.Enabled = False 'Grey out button while running
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor ' Set Cursor to HourGlass

        ' Load and Save Screen Values
        Call LoadScreenValues()
        Call SaveRegSettings()

        ' Open Log File
        blnUseLogFile = True ' Set use Log File flag to True
        strLogFile = GetSetting("QMF2Excel", "Settings", "Log Folder and File", My.Application.Info.DirectoryPath & "\QMF2ExcelLog.txt")
        Try
            objLogFileWriter = New StreamWriter(strLogFile)
        Catch ex As Exception
            blnUseLogFile = False ' If error opening Log File 
        End Try

        ' Clear Log Listbox
        lstLog.Items.Clear()
        Call LogStatus("Submitted")
    End Sub

    Private Sub MainProcess()
        If strScrMode = "Use Screen Values" Then
            Call LogStatus("Using Screen Values")
            Call LoadMainframeFileArrayFromScreenValues()
        Else
            Call LoadMainframeFileArrayFromBatch()
            If File.Exists(strScrBatchFile) = False Then
                Call LogStatus("Error - Could not find Batch File " & strScrBatchFile)
                Exit Sub
            Else
                If blnSuccesfulBatchLaod = False Then
                    Call LogStatus("Error - Could not read Batch File " & strScrBatchFile)
                    Exit Sub
                Else
                    Call LogStatus("Using Batch File")
                End If
            End If
        End If

        Call LoopMainframeFileArry()
    End Sub

    Private Sub LoadMainframeFileArrayFromScreenValues()
        MainframeFileArray(1, 1) = strScrMainframeFile
        MainframeFileArray(1, 2) = strScrExcelFolder & "\" & DateConversion(strScrExcelFile)
        MainframeFileArray(1, 3) = strScrAutoFormat
        intNumberOfMainframeFiles = 1
    End Sub

    Private Sub LoadMainframeFileArrayFromBatch()
        blnSuccesfulBatchLaod = False

        ' Load Batch Valuse in to Array
        Dim xlsApplicationInput As Excel.Application
        Dim xlsWorkBookInput As Excel.Workbook

        ' Open input File
        'Try 
        'xlsApplicationInput = Marshal.GetActiveObject("Excel.Application") 'Grab a running instance of Excel.
        'Catch ex As COMException
        xlsApplicationInput = New Excel.Application 'If no instance exist then create a new one.
        'End Try

        ' Don't Show Excel on Screen
        xlsApplicationInput.Visible = False

        ' Open Batch File (Excel)
        Try
            xlsWorkBookInput = xlsApplicationInput.Workbooks.Open(strScrBatchFile)
        Catch ex As Exception
            ' Close Input Vars
            xlsWorkBookInput = Nothing
            xlsApplicationInput.Quit()
            xlsApplicationInput = Nothing
            Exit Sub
        End Try

        Try
            For intNumberOfMainframeFiles = 1 To 100
                MainframeFileArray(intNumberOfMainframeFiles, 1) = xlsApplicationInput.Worksheets(1).Range("A" & intNumberOfMainframeFiles + 1).Value() ' Mainframe File and Folder
                MainframeFileArray(intNumberOfMainframeFiles, 2) = DateConversion(xlsApplicationInput.Worksheets(1).Range("B" & intNumberOfMainframeFiles + 1).Value()) ' Excel File and Folder
                MainframeFileArray(intNumberOfMainframeFiles, 3) = xlsApplicationInput.Worksheets(1).Range("C" & intNumberOfMainframeFiles + 1).Value() ' Numeric2Text
                If Trim(MainframeFileArray(intNumberOfMainframeFiles, 1)) = "" Then
                    intNumberOfMainframeFiles = intNumberOfMainframeFiles - 1
                    Exit For
                End If
            Next intNumberOfMainframeFiles
        Catch ex As Exception
            ' Close Input Vars
            xlsWorkBookInput.Close()
            xlsWorkBookInput = Nothing
            xlsApplicationInput.Quit()
            xlsApplicationInput = Nothing
            Exit Sub
        End Try

        ' Close Input Vars
        xlsWorkBookInput.Close()
        xlsWorkBookInput = Nothing
        xlsApplicationInput.Quit()
        xlsApplicationInput = Nothing
        blnSuccesfulBatchLaod = True
    End Sub

    Private Sub LoopMainframeFileArry()
        For intCurrentMainframeFile = 1 To intNumberOfMainframeFiles
            Delay(1) ' Slight program pause to make sure everything is caught up
            Call ProcessMainframeFile()
        Next intCurrentMainframeFile
    End Sub

    Private Sub LoadMainframeVarsFromArray()
        ' Load Mainframe Var from Array 
        If IsValidMainframeFile(MainframeFileArray(intCurrentMainframeFile, 1)) Then
            strMainframeFile = MainframeFileArray(intCurrentMainframeFile, 1)
        Else
            strMainframeFile = ""
        End If
    End Sub

    Private Sub LoadExcelVarsFromArray()
        ' Load Excel Vars from Array (Split File name out)
        Try
            strExcelFile = Path.GetFileName(MainframeFileArray(intCurrentMainframeFile, 2))
        Catch ex As Exception
            Call LogStatus("Invalid Excel file name so using Mainframe file name")
            strExcelFile = strMainframeFile & ".xls"
        End Try
        ' Load Excel Vars from Array (Split Folder name out)
        Try
            strExcelFolder = Path.GetDirectoryName(MainframeFileArray(intCurrentMainframeFile, 2))
        Catch ex As Exception
            Call LogStatus("Invalid Excel folder name so using Desktop")
            strExcelFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        End Try

        ' Set Excel Folder
        If strExcelFolder = "" Then
            ' Use Mainframe Folder for Excel Folder
            Call LogStatus("No Excel folder so using Desktop")
            strExcelFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        Else
            ' Check if Excel Folder exist
            If Directory.Exists(strExcelFolder) = False Then
                Call LogStatus("Excel folder does not exist so trying to create it")
                blnSuccessfulTry = True
                Try
                    Directory.CreateDirectory(strExcelFolder)
                Catch ex As Exception
                    ' Use Mainframe Folder for Excel Folder
                    Call LogStatus("Could not create Excel folder so using Desktop")
                    strExcelFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                    blnSuccessfulTry = False
                End Try
                If blnSuccessfulTry = True Then Call LogStatus("New Excel folder was successfully created")
            End If
        End If

        ' Set Excel File
        If strExcelFile = "" Then
            ' Use Mainframe File for Excel File
            Call LogStatus("No Excel file name so using Mainframe file name")
            strExcelFile = Path.GetFileNameWithoutExtension(strMainframeFile) & ".xls"
        End If
        ' Add .xls ext
        If Len(strExcelFile) >= 5 Then
            If Mid(strExcelFile, Len(strExcelFile) - 3, 4) <> ".xls" Then
                strExcelFile = strExcelFile & ".xls"
            End If
        End If

    End Sub

    Private Sub LoadNumeric2TextVarsFromArray()
        If MainframeFileArray(intCurrentMainframeFile, 3) = False Then
            blnNumeric2Text = False
        Else
            blnNumeric2Text = True
        End If
    End Sub

    Private Sub ProcessMainframeFile()
        Try
            Array.Clear(ExcelArray, 0, ExcelArray.Length) ' clear the main array 
            intExcelNumberOfHeadingRows = 0 ' reset to 0
            blnSuccesfulMainframeFileOpen = False ' moved here to better grab file status

            Call LoadMainframeVarsFromArray()
            If strMainframeFile = "" Then
                Call LogStatus("*** Processing Mainframe File (" & intCurrentMainframeFile & " of " & intNumberOfMainframeFiles & ") ***")
                Call LogStatus("Error - Invalid Mainframe File or Folder name")
                Exit Sub
            End If

            Call LogStatus("*** Processing Mainframe File " & strMainframeFile & " (" & intCurrentMainframeFile & " of " & intNumberOfMainframeFiles & ") ***")

            ' Check if Mainframe file name looks valid
            If Not IsValidMainframeFile(strMainframeFile) Then
                Call LogStatus("Error - Mainframe File name not Valid: " & strMainframeFile)
                Exit Sub
            End If

            Call LogStatus("Downloading Mainframe File: " & strMainframeFile)
            ' Delete temp HTML file
            If File.Exists(My.Application.Info.DirectoryPath & "\MainframeFile.htm") = True Then
                Try
                    File.Delete(My.Application.Info.DirectoryPath & "\MainframeFile.htm")
                Catch ex As Exception
                    Call LogStatus("Error: " & ex.ToString)
                    Call LogStatus("Error - Temp HTML file MainframeFile.htm file is open, can not delete")
                    Exit Sub
                End Try
            End If
            ' Download Mainframe file to PC
            If Not DownloadFile(strMainframeFile & ".REPORT", My.Application.Info.DirectoryPath & "\MainframeFile.htm") Then
                Call LogStatus("Error - Could not download: " & strMainframeFile)
                Exit Sub
            End If

            Call LoadExcelVarsFromArray()
            Call LoadNumeric2TextVarsFromArray()

            ' Check if Excel File and Folder exist
            If File.Exists(strExcelFolder & "\" & strExcelFile) = True Then
                If strScrOverwrite = "True" Then
                    ' Overwrite checkbox is set to True, delete the file
                    Call LogStatus("Deleting existing Excel file")
                    Try
                        File.Delete(strExcelFolder & "\" & strExcelFile)
                    Catch
                        Call LogStatus("Error - Excel file is still open, can not delete")
                        Exit Sub
                    End Try
                Else
                    ' Overwrite checkbox is set to False, skip to next file
                    Call LogStatus("Error - Excel file already exist")
                    Exit Sub
                End If
            End If

            Call OpenMainframeFile()
            If blnSuccesfulMainframeFileOpen = False Then
                Call LogStatus("Error - Could not read Temp HTML file MainframeFile.htm")
                Exit Sub
            End If
            Call ReadMainframeFile()
        Catch ex As Exception
            Call LogStatus("Error: " & ex.ToString)
        Finally
            If blnSuccesfulMainframeFileOpen = True Then Call CloseMainframeFile()
        End Try
    End Sub

    Private Sub OpenMainframeFile()
        Call LogStatus("Opening Mainframe (HTML) File")
        Try
            sr = New StreamReader(My.Application.Info.DirectoryPath & "\MainframeFile.htm")
        Catch ex As Exception
            Exit Sub
        End Try
        blnSuccesfulMainframeFileOpen = True
    End Sub

    Private Sub ReadMainframeFile()
        ' Skip first 8 lines (HTML Stuff)
        Call LogStatus("Check if file is in Mainframe HTML Foramt")
        For intLineNumber = 1 To 8
            strLine = sr.ReadLine()
            ' Check if file ends too soon
            If sr.EndOfStream = True Then
                Call LogStatus("Error - Mainframe File Ended")
                Exit Sub
            End If
            ' Check if line 8 is <PRE> (something QMF uses, but not a lot of other programs use)
            If intLineNumber = 8 And strLine <> "<PRE>" Then
                Call LogStatus("Error - Mainframe File is not in Mainframe HTML Format")
                Exit Sub
            End If
        Next intLineNumber

        ' Grab Headings 
        Call LogStatus("Grabing Mainframe Headings")
        blnHeadingsFound = False
        For intLineNumber = 9 To 20
            strLine = sr.ReadLine()
            ' Check if file ends too soon
            If sr.EndOfStream = True Then
                Call LogStatus("Error - Mainframe File Ended")
                Exit Sub
            End If

            intNumberHeadingRows = intLineNumber - 8
            RawHeadingArray(intNumberHeadingRows) = strLine
            ' Looking for two lines in a row with just one space to signify the end of headings and the beginning of columns
            If strLine = " " And RawHeadingArray(intNumberHeadingRows - 1) = " " Then
                intNumberHeadingRows = intNumberHeadingRows - 2
                blnHeadingsFound = True
                Exit For
            End If
            ' Seems some Mainframe files have two blank lines instead 
            If strLine = Nothing And RawHeadingArray(intNumberHeadingRows - 1) = Nothing Then
                intNumberHeadingRows = intNumberHeadingRows - 2
                blnHeadingsFound = True
                Exit For
            End If
        Next intLineNumber
        If blnHeadingsFound = False Then
            Call LogStatus("Error - Headings in Mainframe File Not Found")
            Exit Sub
        End If

        ' Grab Raw Column Names (and Divider Lines)
        Call LogStatus("Grabing Raw Mainframe Column Names")
        strDiv = ""
        For intCnt = 1 To 10
            intLineNumber = intLineNumber + 1
            strLine = sr.ReadLine()
            ' Check if file ends too soon
            If sr.EndOfStream = True Then
                Call LogStatus("Error - Mainframe File Ended")
                Exit Sub
            End If

            intNumberRawColumnNameRows = intCnt
            RawColumnNameArray(intNumberRawColumnNameRows) = strLine

            If Mid(strLine, 3, 1) = "-" Then
                intNumberRawColumnNameRows = intNumberRawColumnNameRows - 1 ' Roll counter back to last line with actaul column name
                strDiv = strLine ' Populate Div Line      
                Exit For
            End If
        Next
        If strDiv = "" Then
            Call LogStatus("Error - Column Names in Mainframe File Not Found")
            Exit Sub
        End If

        ' Get Column Name Lenghts
        Call LogStatus("Grabing Mainframe Column Lenghts")
        intNumberColumns = 1
        flgDivStatus = "S"

        For intCnt = 3 To Len(strDiv)
            If Mid(strDiv, intCnt, 1) = "-" Then
                If flgDivStatus = "S" Then DivArray(intNumberColumns, 1) = intCnt : flgDivStatus = "E"
            Else
                If flgDivStatus = "E" Then DivArray(intNumberColumns, 2) = intCnt - 1 : flgDivStatus = "S" : intNumberColumns = intNumberColumns + 1
            End If
        Next intCnt
        DivArray(intNumberColumns, 2) = intCnt

        ' Get Column Names
        Call LogStatus("Grabing Mainframe Column Names")
        For intCnt = 1 To intNumberColumns
            ColumnNames(intCnt) = "" ' Clear Names
            For intCnt2 = 1 To intNumberRawColumnNameRows
                ColumnNames(intCnt) = ColumnNames(intCnt) & " " & Trim(Mid(RawColumnNameArray(intCnt2), DivArray(intCnt, 1), DivArray(intCnt, 2) - DivArray(intCnt, 1) + 1))
            Next intCnt2
            ColumnNames(intCnt) = Trim(ColumnNames(intCnt))
        Next intCnt

        ' Get Max Char Length of Column 1
        intMaxCol1Chars = Len(ColumnNames(1))

        ' Put Headings into main Array
        For intOutputRow = 0 To intNumberHeadingRows - 1
            ExcelArray(intOutputRow, 0) = DateConversion(Trim(RawHeadingArray(intOutputRow + 1)))
            intExcelNumberOfHeadingRows = intOutputRow + 1
        Next intOutputRow

        If intNumberHeadingRows > 1 Then intOutputRow = intOutputRow + 1 ' Add a blank line between Headings and Columns

        blnAnyTotals = False ' jjb-total
        For intCnt = 0 To 100 ' jjb-total
            ColumnTotals(intCnt) = "No" ' jjb-total
        Next intCnt ' jjb-total

        ' Put Columnn Names into main Array
        For intCnt = 0 To intNumberColumns - 1

            ' Check if last Column Name character is a "*", if so remove it and mark as Total col ' jjb-total
            If Mid(ColumnNames(intCnt + 1), Len(ColumnNames(intCnt + 1)), 1) = "*" And blnNumeric2Text = True Then ' jjb-total
                ExcelArray(intOutputRow, intCnt) = Mid(ColumnNames(intCnt + 1), 1, Len(ColumnNames(intCnt + 1)) - 1) ' jjb-total
                ColumnTotals(intCnt + 1) = "Yes" ' jjb-total
                blnAnyTotals = True ' jjb-total
            Else ' jjb-total
                ExcelArray(intOutputRow, intCnt) = ColumnNames(intCnt + 1)
                ColumnTotals(intCnt + 1) = "No" ' jjb-total
            End If ' jjb-total

            intExcelColumnNameRow = intOutputRow + 1
        Next intCnt
        intOutputRow = intOutputRow + 1

        intFirstExcelDataRow = intOutputRow + 1 ' jjb-total

        ' Put Mainframe data items into main Array
        Call LogStatus("Grabing Mainframe Data")
        For intOutputRow = intOutputRow To 65535
            If intOutputRow Mod 10 Then Application.DoEvents() ' every 10th row / keep app responsive

            strPreviousLine = strLine
            strLine = sr.ReadLine()
            ' Check for early file end
            If sr.EndOfStream = True Then
                Call LogStatus("Error - Mainframe File Ended")
                Exit Sub
            End If

            ' Check if end of Mainframe data
            If strLine = " " And strPreviousLine = " " Then
                intOutputRow = intOutputRow - 1 ' Would be -2 to roll back to last row with data, but we want one line blank between data and footer anyway
                blnFootersFound = True
                Exit For
            End If
            ' Some Mainframe files use two blank lines
            If strLine = Nothing And strPreviousLine = Nothing Then
                intOutputRow = intOutputRow - 1 ' Would be -2 to roll back to last row with data, but we want one line blank between data and footer anyway
                blnFootersFound = True
                Exit For
            End If
            ' Move each column to main Array
            For intCnt = 0 To intNumberColumns - 1
                strDataItem = Trim(Mid(strLine, DivArray(intCnt + 1, 1), DivArray(intCnt + 1, 2) - DivArray(intCnt + 1, 1) + 1))

                ' Get Max Char Length of Column 1
                If intCnt = 0 And Len(strDataItem) > intMaxCol1Chars Then intMaxCol1Chars = Len(strDataItem)
                ' Make sure Total lines (====) don't try to become formulas
                If Mid(strDataItem, 1, 1) = "=" Then
                    If strScrSkipTotals = True Then
                        strLine = sr.ReadLine() ' skip the actual number line (the next line)
                        intOutputRow = intOutputRow - 1 ' since skipping this line, don't count it
                        Exit For ' skip the "====" line (this line)
                    Else
                        ExcelArray(intOutputRow, intCnt) = "'" & strDataItem
                    End If
                Else
                    ' ** Auto Format Code Begin **
                    If blnNumeric2Text = True Then
                        If IsNumeric(strDataItem) Then
                            If IsDecimal(strDataItem) Then
                                ExcelArray(intOutputRow, intCnt) = FormatCurrency(strDataItem, , , False)
                            Else
                                If Mid(ColumnNames(intCnt + 1), 1, 3) = "SUM" Or Mid(ColumnNames(intCnt + 1), 1, 5) = "COUNT" Or Mid(ColumnNames(intCnt + 1), 1, 7) = "AVERAGE" Or Mid(ColumnNames(intCnt + 1), 1, 3) = "CAL" Or Mid(ColumnNames(intCnt + 1), 1, 3) = "NUM" Or Mid(ColumnNames(intCnt + 1), 1, 4) = "YEAR" Or Mid(ColumnNames(intCnt + 1), 1, 5) = "MONTH" Then ' make clearly numeric stuff... nurmeric
                                    ExcelArray(intOutputRow, intCnt) = strDataItem
                                Else
                                    ExcelArray(intOutputRow, intCnt) = CStr("'" & strDataItem)
                                End If
                            End If
                        Else
                            If Trim(strDataItem) = "-" Then ' Remove cells with "-", usualy caused by left outer joins
                                ExcelArray(intOutputRow, intCnt) = ""
                            Else
                                ExcelArray(intOutputRow, intCnt) = strDataItem
                            End If
                        End If
                        ' ** Auto Format Code End **
                    Else
                        ExcelArray(intOutputRow, intCnt) = strDataItem
                    End If
                End If
            Next intCnt
        Next intOutputRow

        ' Insert Totals ' jjb-total (whole section)
        If blnNumeric2Text = True Then
            If blnAnyTotals = True Then
                Call LogStatus("Insert Excel Totals")

                For intCnt = 1 To 100
                    If ColumnTotals(intCnt) = "Yes" Then
                        ExcelArray(intOutputRow, intCnt - 1) = "=SUM(" & ColumnLetter(intCnt) & intFirstExcelDataRow & ":" & ColumnLetter(intCnt) & intOutputRow & ")"
                    End If
                Next intCnt

                intOutputRow = intOutputRow + 1
                intTotalRow = intOutputRow
            End If
        End If

        ' Grab Footers 
        Call LogStatus("Grabing Mainframe Footers")
        strEndFound = "N"
        For intCnt = 1 To 10

            strLine = sr.ReadLine()

            ' Check if file ends too soon
            If sr.EndOfStream = True Then
                Exit For
            End If

            If Trim(strLine) = "</PRE>" Then
                strEndFound = "Y"
                Exit For
            End If

            intOutputRow = intOutputRow + 1
            If intOutputRow > 65535 Then
                strEndFound = "TooLong"
                Exit For
            End If
            ExcelArray(intOutputRow, 0) = Trim(strLine)
        Next intCnt

        If strEndFound = "TooLong" Then Call LogStatus("Mainframe File Contains More than 65536 Rows")
        If strEndFound = "N" Then
            Call LogStatus("Error - Mainframe File Ended (Trying to Process)")
        Else
            Call LogStatus("Mainframe File Completely Read")
        End If

        ' Write Main Array to new Excel file
        Call LogStatus("Opening Excel File")
        Call OpenExcelFile()

        Call LogStatus("Write Data to Excel File")
        ' Write Excel data in the background
        blnExcelCopyComplete = False
        Dim t As System.Threading.Thread
        t = New System.Threading.Thread(AddressOf Me.BackgroundProcess)
        t.Start()
        Do While blnExcelCopyComplete = False
            Delay(1)
        Loop

        If blnNumeric2Text = True Then
            Call LogStatus("Formatting Excel File")

            ' Report Headings
            If intExcelNumberOfHeadingRows > 0 Then
                With xlsWorkSheet.Range("A1", "A" & intExcelNumberOfHeadingRows)
                    .Font.Name = "Arial"
                    .Font.Bold = True
                    .Font.Size = 10
                End With
            End If

            ' Column Headings
            With xlsWorkSheet.Range("A" & intExcelColumnNameRow, ColumnLetter(intNumberColumns) & intExcelColumnNameRow)
                .Font.Bold = True
                .Font.Name = "Arial"
                .Font.ColorIndex = 2 ' White
                .Font.Size = 9
                .HorizontalAlignment = -4108 ' Center
                .Cells.WrapText = True
                .Cells.Interior.Colorindex = 11 ' Dark Blue
            End With

            ' Report Data
            With xlsWorkSheet.Range("A" & intExcelColumnNameRow + 1, ColumnLetter(intNumberColumns) & intOutputRow)
                .Font.Name = "Arial"
                .Font.Size = 9
            End With

            ' Report Totals ' jjb-total
            If blnAnyTotals = True Then
                For intCnt = 1 To 100
                    If ColumnTotals(intCnt) = "Yes" Then
                        With xlsWorkSheet.Range(ColumnLetter(intCnt) & intTotalRow, ColumnLetter(intCnt) & intTotalRow)
                            .Font.Name = "Arial"
                            .Font.Size = 9
                            .Font.Bold = True
                            .Borders(8).LineStyle = Excel.XlLineStyle.xlDouble
                            .Borders(8).Weight = Excel.XlBorderWeight.xlThick
                        End With
                    End If
                Next intCnt
            End If
        End If

        ' Report Data Auto Fit
        With xlsWorkSheet.Range("A" & intExcelColumnNameRow + 1, ColumnLetter(intNumberColumns) & intOutputRow)
            .EntireColumn.AutoFit()
        End With

        ' Space first column correctly
        If intNumberHeadingRows >= 1 Then xlsWorkSheet.Columns(1).ColumnWidth = intMaxCol1Chars * 1.4

        Call LogStatus("Close Excel File")
        Call CloseExcelFile()
    End Sub
    Sub OpenExcelFile()
        xlsApplication = New Excel.Application 'If no instance exist then create a new one.

        xlsApplication.Visible = False ' Set to see magic in action
        xlsWorkBook = xlsApplication.Workbooks.add

        ' Get a new workbook.
        xlsWorkSheet = xlsWorkBook.activeSheet
    End Sub

    Private Sub BackgroundProcess()
        xlRng = xlsWorkSheet.Range("A1:" & ColumnLetter(intNumberColumns) & intOutputRow + 1)
        xlRng.Value = ExcelArray
        blnExcelCopyComplete = True
    End Sub

    Sub CloseExcelFile()
        Try
            xlsWorkBook.SaveAs(strExcelFolder & "\" & strExcelFile)
        Catch ex As Exception
            Call LogStatus("Error Saving Excel File")
        Finally
            xlsApplication.Quit()
            xlsApplication = Nothing
            xlsWorkBook = Nothing
            xlsWorkSheet = Nothing
        End Try
    End Sub

    Private Sub CloseMainframeFile()
        Call LogStatus("Close Mainframe File")
        sr.Close()
    End Sub

    Private Sub SubmitComplete()
        ' Clean Up
        Call LogStatus("Complete")
        If blnUseLogFile = True Then objLogFileWriter.Close() ' Close Log File
        Me.Cursor = System.Windows.Forms.Cursors.Default ' Set Cursor to HourGlass
        btnSubmit.Enabled = True 'Re-enable button
    End Sub

    Private Sub LogStatus(ByVal Message As String)
        If blnUseLogFile = True Then objLogFileWriter.WriteLine(Now & " - " & Message) ' Write to Log File
        lstLog.Items.Add(Now & " - " & Message) ' Display on Screen
        lstLog.SelectedIndex = lstLog.Items.Count - 1 ' Current mesage is always selected
        Application.DoEvents() ' keep app responsive
    End Sub

    Private Sub Delay(ByVal DelayInSeconds As Integer)
        Dim ts As TimeSpan
        Dim targetTime As DateTime = DateTime.Now.AddSeconds(DelayInSeconds)
        Do
            ts = targetTime.Subtract(DateTime.Now)
            Application.DoEvents() ' keep app responsive
            System.Threading.Thread.Sleep(50) ' reduce CPU usage
        Loop While ts.TotalSeconds > 0
    End Sub

    Function IsDecimal(ByVal psValue As String) As Boolean
        If IsNumeric(psValue) And InStr(psValue, ".") Then
            Return True
        Else
            Return False
        End If
    End Function

    Function IsValidMainframeFile(ByVal psValue As String) As Boolean
        Dim intCounter
        IsValidMainframeFile = False
        ' Change 46 to 47 if "." is allowed
        For intCounter = 1 To Len(psValue)
            If Asc(Mid(psValue, intCounter, 1)) <= 31 Then Exit Function
            If Asc(Mid(psValue, intCounter, 1)) >= 33 And Asc(Mid(psValue, intCounter, 1)) <= 39 Then Exit Function
            If Asc(Mid(psValue, intCounter, 1)) >= 42 And Asc(Mid(psValue, intCounter, 1)) <= 44 Then Exit Function
            If Asc(Mid(psValue, intCounter, 1)) >= 46 And Asc(Mid(psValue, intCounter, 1)) <= 46 Then Exit Function
            If Asc(Mid(psValue, intCounter, 1)) >= 58 And Asc(Mid(psValue, intCounter, 1)) <= 64 Then Exit Function
            If Asc(Mid(psValue, intCounter, 1)) >= 91 And Asc(Mid(psValue, intCounter, 1)) <= 94 Then Exit Function
            If Asc(Mid(psValue, intCounter, 1)) = 96 Then Exit Function
            If Asc(Mid(psValue, intCounter, 1)) >= 123 Then Exit Function
        Next intCounter
        IsValidMainframeFile = True
    End Function

    Function DownloadFile(ByVal psMainframeFile As String, ByVal psPCFile As String) As Boolean
        DownloadFile = False
        Dim ff As clsFTP

        Try
            ff = New clsFTP
            ' Setup the appropriate properties.
            ff.RemoteHost = strScrMainframeHost
            ff.RemoteUser = strScrMainframeUserId
            ff.RemotePassword = strScrMainframePassword
            ' Attempt to log into the FTP Server.
            If (ff.Login()) Then
                ' Download a file.
                ff.SetBinaryMode(False)
                ''Call RefreshScreen() 
                ff.DownloadFile(psMainframeFile, psPCFile)
                DownloadFile = True
            End If
            '        Catch ex As Exception
            '           Call LogStatus("Error: " & ex.ToString) ' jjb
        Finally
            ff.CloseConnection()
        End Try
    End Function

    Function ColumnLetter(ByVal ColumnNumber As Integer) As String
        If ColumnNumber > 26 Then

            ' 1st character:  Subtract 1 to map the characters to 0-25,
            '                 but you don't have to remap back to 1-26
            '                 after the 'Int' operation since columns
            '                 1-26 have no prefix letter

            ' 2nd character:  Subtract 1 to map the characters to 0-25,
            '                 but then must remap back to 1-26 after
            '                 the 'Mod' operation by adding 1 back in
            '                 (included in the '65')

            ColumnLetter = Chr(Int((ColumnNumber - 1) / 26) + 64) & Chr(((ColumnNumber - 1) Mod 26) + 65)
        Else
            ' Columns A-Z
            ColumnLetter = Chr(ColumnNumber + 64)
        End If
    End Function

    Function DateConversion(ByVal psString As String)
        If InStr(psString, "{") = 0 Then
            DateConversion = psString
            Exit Function
        End If

        If InStr(psString, "}") = 0 Then
            DateConversion = psString
            Exit Function
        End If

        Dim i As Integer
        Dim DateCode As String
        Dim dteUseDate As Date
        Dim intOffSet As Integer
        Dim DateFormat As String
        DateConversion = ""

        For i = 1 To Len(psString)
            If Mid(psString, i, 1) = "{" Then
                DateCode = Mid(psString, i + 1, 4)
                intOffSet = -1
                Select Case DateCode
                    Case "TODY"
                        dteUseDate = Now
                    Case "FDCM"
                        dteUseDate = DateSerial(Year(Now), Month(Now), 1)
                    Case "LDCM"
                        dteUseDate = DateSerial(Year(Now), Month(Now) + 1, 1 - 1)
                    Case "FDPM"
                        dteUseDate = DateSerial(Year(Now), Month(Now) - 1, 1)
                    Case "LDPM"
                        dteUseDate = DateSerial(Year(Now), Month(Now), 1 - 1)
                    Case "PSUN"
                        Do
                            dteUseDate = DateTime.Today.AddDays(intOffSet)
                            If dteUseDate.DayOfWeek = DayOfWeek.Sunday Then Exit Do
                            intOffSet = intOffSet - 1
                        Loop
                    Case "PMON"
                        Do
                            dteUseDate = DateTime.Today.AddDays(intOffSet)
                            If dteUseDate.DayOfWeek = DayOfWeek.Monday Then Exit Do
                            intOffSet = intOffSet - 1
                        Loop
                    Case "PTUE"
                        Do
                            dteUseDate = DateTime.Today.AddDays(intOffSet)
                            If dteUseDate.DayOfWeek = DayOfWeek.Tuesday Then Exit Do
                            intOffSet = intOffSet - 1
                        Loop
                    Case "PWED"
                        Do
                            dteUseDate = DateTime.Today.AddDays(intOffSet)
                            If dteUseDate.DayOfWeek = DayOfWeek.Wednesday Then Exit Do
                            intOffSet = intOffSet - 1
                        Loop
                    Case "PTHU"
                        Do
                            dteUseDate = DateTime.Today.AddDays(intOffSet)
                            If dteUseDate.DayOfWeek = DayOfWeek.Thursday Then Exit Do
                            intOffSet = intOffSet - 1
                        Loop
                    Case "PFRI"
                        Do
                            dteUseDate = DateTime.Today.AddDays(intOffSet)
                            If dteUseDate.DayOfWeek = DayOfWeek.Friday Then Exit Do
                            intOffSet = intOffSet - 1
                        Loop
                    Case "PSAT"
                        Do
                            dteUseDate = DateTime.Today.AddDays(intOffSet)
                            If dteUseDate.DayOfWeek = DayOfWeek.Saturday Then Exit Do
                            intOffSet = intOffSet - 1
                        Loop
                    Case Else
                        ' Error
                        DateConversion = psString
                        Exit Function
                End Select

                DateFormat = ""
                ' Look for end bracket }, and get Date Format
                i = i + 7
                Do
                    ' Error 
                    If i > Len(psString) Then
                        DateConversion = psString
                        Exit Function
                    End If
                    If Mid(psString, i, 1) = "'" Then
                        i = i + 1
                        Exit Do
                    End If
                    DateFormat = DateFormat & Mid(psString, i, 1)
                    i = i + 1
                Loop

                DateFormat = DateFormat.Replace("D", "d")
                DateFormat = DateFormat.Replace("m", "M")
                DateFormat = DateFormat.Replace("Y", "y")
                DateFormat = DateFormat.Replace("H", "h")
                DateFormat = DateFormat.Replace("N", "m")
                DateFormat = DateFormat.Replace("S", "s")
                DateFormat = DateFormat.Replace("T", "t")
                Try
                    If DateFormat = "Q" Or DateFormat = "q" Then
                        Select Case Month(dteUseDate)
                            Case 1 To 3
                                DateConversion = DateConversion & "1st"
                            Case 4 To 6
                                DateConversion = DateConversion & "2nd"
                            Case 7 To 9
                                DateConversion = DateConversion & "3rd"
                            Case Else
                                DateConversion = DateConversion & "4th"
                        End Select
                    Else
                        DateConversion = DateConversion & Format(dteUseDate, DateFormat)
                    End If
                Catch
                    ' do nothing for now
                End Try
            Else
                DateConversion = DateConversion & Mid(psString, i, 1)
            End If
        Next i

    End Function

End Class
