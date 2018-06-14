Imports System.IO
Imports System.Data.OleDb
Imports MySql.Data.MySqlClient
Imports System.Windows.Forms
Imports System.Configuration
Imports System.Text
Imports Microsoft.Office.Interop

Public Class Form2
    Dim strSql As String
    Dim CountSteps As Integer
    Dim connStrBMG As String = "Database=belo_database;Data Source= localhost;Port=3308;User Id=root ;Password='';UseCompression=True;Convert Zero Datetime=True"
Private Sub Form2_Load( sender As Object,  e As EventArgs) Handles MyBase.Load

End Sub

Private Sub Button1_Click( sender As Object,  e As EventArgs) Handles Button1.Click
        OpenFileDialog1.Title = "Select CSV File"
        OpenFileDialog1.InitialDirectory= Application.StartupPath()
        OpenFileDialog1.Filter="csv (*.csv)|(*.csv)|All Files (*.*)|(*.*)"
        OpenFileDialog1.FileName = "*.csv"
        Dim OpenFileName As String
        Dim result As DialogResult = OpenFileDialog1.ShowDialog()
        
        Try
            'If Pressed OK
            If (result = DialogResult.OK) then
            OpenFileName = OpenFileDialog1.FileName
            txtFileDirectory.Text = OpenFileName
            'cmdDisplay_Click(Me,EventArgs.Empty)
        'If Pressed Cancel
            ElseIf (result = DialogResult.Cancel) Then
                Return
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    
End Sub
    
Private Sub cmdDisplay_Click( sender As Object,  e As EventArgs) Handles cmdDisplay.Click
        If txtFileDirectory.Text <> "" then
            Dim FilePath As String = txtFileDirectory.Text 
            Dim fi As New FileInfo(FilePath)
            Dim connectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=Text;Data Source=" & fi.DirectoryName

            Dim conn As New OleDbConnection(connectionString)
            conn.Open()

            'the SELECT statement is important here, 
            'and requires some formatting to pull dates and deal with headers with spaces.
Dist:
            Dim cmdSelect As New OleDbCommand("SELECT TOP 200 Accountno, Memberdate,Firstname, Lastname,Pointsearned FROM " & fi.Name, conn)

            Dim adapter1 As New OleDbDataAdapter
            adapter1.SelectCommand = cmdSelect
            Dim ds As New DataSet
            Try
                adapter1.Fill(ds, "DATA")
                DataGridView1.DataSource = ds.Tables(0).DefaultView 
                'DataGridView1.DataBindings
                conn.Close()
                ProgBar.Maximum = DataGridview1.RowCount
                'cmdCheck_Click(Me,EventArgs.Empty)
            Catch ex As Exception
                MsgBox(ex.Message)
                GoTo Dist
            End Try
            
        Else
            MsgBox("Nothing to Display!",vbExclamation + vbOKOnly,"No Records found!")
        End If 
End Sub

Private Sub cmdCheck_Click( sender As Object,  e As EventArgs) Handles cmdCheck.Click
        Dim xCard As String
        Dim xPoints As Integer
        Dim x As Integer
        Dim Total As Integer = DataGridView1.RowCount - 1
        Try 
            For i As Integer = 0 to Total
            ProgBar.Value = ProgBar.Value + 1
            xCard = DataGridView1.Rows(i).Cells("Accountno").Value.ToString
                If Mid(xCard,1,2) = "WV" then
                    xCard = Replace(xCard,"WV","",1) 
                End If
            xPoints = DataGridView1.Rows(i).Cells("Pointsearned").Value
            ReadChecker(xCard,xPoints,i)
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        MsgBox("Done")
End Sub

Private Sub ReadChecker(ByVal CardNum As string, ByVal TotalPoints As Integer, ByVal steps As Integer ) 
        Dim sr As New StreamReader(Application.StartupPath & "\SqlConnection.dll")
        Dim query As String = "select * from `Loyalty_Card_Info_Copy` where CardNo='" & CardNum & "'"
        Dim connStrSql As String = sr.ReadLine
        Dim connection1 As New MySqlConnection(connStrSql)
        Dim connection As New MySqlConnection(connStrBMG)
        Dim cmd As new MySqlCommand(query,connection)
        Dim reader As MySqlDataReader
        connection.Open
        Try
            reader = cmd.ExecuteReader
            If reader.HasRows then
                While reader.Read
                    If TotalPoints <> Val(reader.Item("total_points").ToString) then
                        If txtAccountno.Text = "" then
                            txtAccountno.Text = DataGridView1.Rows(steps).Cells("AccountNo").Value.ToString & "|"
                            CountSteps = CountSteps + 1
                        Else
                            txtAccountno.Text = txtAccountno.Text & DataGridView1.Rows(steps).Cells("AccountNo").Value.ToString & "|"
                            CountSteps = CountSteps + 1
                        End If
                    Else
                        'Ignore
                    End If
                End While
            End If
        Catch ex As Exception
            Dim err As String
            err = ex.Message
            If InStr(1,ex.Message,"Object reference not set to an instance of an object.") then
                
            else
                MsgBox(ex.Message)
            End If
        End Try
        connection.Close
End Sub

Private Sub cmdExcel_Click( sender As Object,  e As EventArgs) Handles cmdExcel.Click
        Dim stepx As Integer = Val(Label2.Text)
        Dim xAccountNo As String() = txtAccountno.Text.Split("|") 
        Dim xls As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
        Dim deto As String 
        deto = now().ToString("yyyyMMddHH")
        If xls Is Nothing Then
            MessageBox.Show("Excel is not properly installed!!")
            Return
        End If
    
        Dim book As Excel.Workbook
        Dim sheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value

        xls.Workbooks.Add()
        book = xls.ActiveWorkbook
        sheet = book.ActiveSheet


        'ExcelSheet Content
        sheet.Cells(1,1) = "Account No." : sheet.Cells(1,1).FONT.NAME = "Comic Sans"
        sheet.Cells(1,2) = "Member Date" 
        sheet.Cells(1,3) = "Firstname"
        sheet.Cells(1,4) = "Lastname"
        sheet.Cells(1,5) = "Total Points"
        '
        '
        For i As Integer = 0 to CountSteps
            stepx = i + 2
        sheet.Cells(stepx,1) = xAccountNo(i)
        Next

        


        xls.Visible = True
        releaseObject(sheet)
        releaseObject(book)
        releaseObject(xls)
        'Save Sheet
        'book.SaveAs(CurDir() & "\Results " & deto & ".xls",Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, _
        'Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
End Sub

Private Sub releaseObject(ByVal obj As Object)
        'Release an automation object
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
End Sub

Private Sub ProgBar_Click( sender As Object,  e As EventArgs) Handles ProgBar.Click

End Sub
End Class