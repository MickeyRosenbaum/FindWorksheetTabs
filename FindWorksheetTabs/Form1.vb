Imports System.IO
Imports Microsoft.Office.Interop
Public Class Form1

    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet

    Private Sub btnFindTabs_Click(sender As Object, e As EventArgs) Handles btnFindTabs.Click
        Dim fn As String = OpenExcelFile()
        Dim dt As DataTable = GetWorksheetTabs(fn)
        CloseExcel()
        DataGridView1.DataSource = dt
        DataGridView1.Columns(0).Width = 80
        DataGridView1.Columns(1).Width = 300

    End Sub

    Public Function OpenExcelFile() As String
        'try to open an excel file
        Dim FileOpened As Boolean = False

        'see if excel is running.  if it is, ask to close
        If CheckForExcelOpen() Then
            Dim ans As MsgBoxResult = MsgBox("There is a copy of Excel running.  I need all copies closed.  May I close Excel?", vbYesNo, "Excel Running")
            If ans = vbYes Then 'ok to close
                CloseExcel()
            Else    'not ok to close, tell user file not opened
                Return FileOpened
                Exit Function
            End If
        End If

        'open copy of excel
        xlApp = New Excel.Application

        'some variables
        Dim FileName As String = ""
        Dim WorkSheetName As String = ""

        'set up dialog box
        OpenFileDialog1.Title = "Open Excel Files"
        OpenFileDialog1.Filter = "Excel Files (*.xls*)|*.xls*"  'show Excel files

        Dim UserPressedCancel As DialogResult = OpenFileDialog1.ShowDialog()    'open the file selection dialog box

        'user cancelled, close excel and tell user
        If UserPressedCancel = DialogResult.Cancel Then
            CloseExcel()
            Return ""
            Exit Function
        End If

        'see if file already exists
        If Dir(OpenFileDialog1.FileName) = "" Then            'if the file name does not exist yet
            MsgBox("You must select a file")
            CloseExcel()
            Exit Function
        Else  'the file name already exists
            FileName = OpenFileDialog1.FileName

        End If

        Return FileName

    End Function

    Public Function CloseExcelFile() As Boolean
        'put in try loop
        Try
            xlApp.ActiveWorkbook.Save()               'save the workbook
        Catch ex As Exception
        Finally
            CloseExcel()
        End Try

        'tell user file has been written
        If OpenFileDialog1.FileName <> "" Then
            MsgBox(OpenFileDialog1.FileName & " has been written.", vbOKOnly, "File Written")
        End If
        Return True

    End Function

    Function GetWorksheetTabs(ByVal filename As String) As DataTable

        'see what worksheet tabs alread exist in the excel worksheet
        xlBook = xlApp.Workbooks.Open(filename)             'get the file name selected

        Dim intSheets As Integer = xlApp.Worksheets.Count      'how many sheets are there?

        'define a crlf string
        Dim S As String = vbCrLf

        Dim dt As DataTable = New DataTable
        dt.Columns.Add("Tab No", Type.GetType("System.String"))
        dt.Columns.Add("Tab Name", Type.GetType("System.String"))



        For i As Integer = 1 To intSheets
            dt.Rows.Add(i - 1)
            dt.Rows(i - 1)(0) = i
            dt.Rows(i - 1)(1) = xlApp.Worksheets(i).Name   'add in the worksheet name
        Next i

        GetWorksheetTabs = dt
    End Function
    Private Sub CloseExcel()
        If Not IsNothing(xlApp) Then    'if excel is running
            If xlApp.Workbooks.Count > 0 Then   'and there is a workbook open
                xlApp.Workbooks.Close()         'close the workbook
            End If

            'quit excel
            xlApp.Application.Quit()
            xlApp = Nothing
        End If

        'make sure that there are no excel processes running
        Dim proc As System.Diagnostics.Process
        For Each proc In System.Diagnostics.Process.GetProcessesByName("EXCEL")
            proc.Kill()
        Next

    End Sub

    Private Function CheckForExcelOpen() As Boolean
        'see if there are any processes named excel running
        Dim proc As System.Diagnostics.Process
        For Each proc In System.Diagnostics.Process.GetProcessesByName("EXCEL")
            Return True
        Next
        Return False
    End Function

    Private Sub btnWriteToNotepad_Click(sender As Object, e As EventArgs) Handles btnWriteToNotepad.Click
        'copy the datagridview to the clipboard
        CopyDataGridViewToClipboard(DataGridView1)

        'open notepad and wait for it to completely load
        Dim p As System.Diagnostics.Process = System.Diagnostics.Process.Start("notepad.exe")
        p.WaitForInputIdle()

        ' paste the data from the clipboard
        SendKeys.Send("^V")

    End Sub

    Private Sub CopyDataGridViewToClipboard(ByRef dgv As DataGridView)
        Dim s As String = ""
        Dim oCurrentCol As DataGridViewColumn    'Get header
        oCurrentCol = dgv.Columns.GetFirstColumn(DataGridViewElementStates.Visible)
        Do
            s &= oCurrentCol.HeaderText & Chr(Keys.Tab)
            oCurrentCol = dgv.Columns.GetNextColumn(oCurrentCol,
               DataGridViewElementStates.Visible, DataGridViewElementStates.None)
        Loop Until oCurrentCol Is Nothing
        s = s.Substring(0, s.Length - 1)
        s &= Environment.NewLine    'Get rows
        For Each row As DataGridViewRow In dgv.Rows
            oCurrentCol = dgv.Columns.GetFirstColumn(DataGridViewElementStates.Visible)
            Do
                If row.Cells(oCurrentCol.Index).Value IsNot Nothing Then
                    s &= row.Cells(oCurrentCol.Index).Value.ToString
                End If
                s &= Chr(Keys.Tab)
                oCurrentCol = dgv.Columns.GetNextColumn(oCurrentCol,
                      DataGridViewElementStates.Visible, DataGridViewElementStates.None)
            Loop Until oCurrentCol Is Nothing
            s = s.Substring(0, s.Length - 1)
            s &= Environment.NewLine
        Next    'Put to clipboard
        Dim o As New DataObject
        o.SetText(s)
        Clipboard.SetDataObject(o, True)
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        If System.Diagnostics.Debugger.IsAttached = False Then
            Me.Text = "Find Worksheet Tabs - Version " & "Version : " &
            My.Application.Deployment.CurrentVersion.ToString
        Else
            Me.Text = "Find Worksheet Tabs - Version " & "Debug Mode:" & My.Application.Info.Version.ToString
        End If

    End Sub
End Class
