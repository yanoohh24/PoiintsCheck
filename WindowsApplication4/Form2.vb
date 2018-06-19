Imports System.IO
Imports System.Data.OleDb
Imports MySql.Data.MySqlClient
Imports System.Windows.Forms
Imports System.Configuration
Imports System.Text
Imports Microsoft.Office.Interop
Imports System.IO.Ports
Imports System.IO.Compression
Imports System.ComponentModel
Imports System.ComponentModel.Design.Serialization
Imports System.Threading
Imports System.Reflection
Imports System.ServiceProcess
Imports System.Management
Imports System.Text.RegularExpressions

Public Class Form2
    Dim strSql ,fname, lname , xRec ,xTotal ,Points ,AccMain ,xTop As String 
    Dim Nofname, nolname, nofullname, noRec, noTotal, NoPoints As String
    Dim fullname = ""
    Dim CountSteps As Integer
    Dim dt As Date
    Dim neg As Boolean = False : Dim negDB as Boolean = False
    Dim connStrLocal As String = "Database=belo_database;Data Source= localhost;Port=3308;User Id=root ;Password='';UseCompression=True;Convert Zero Datetime=True"
    Dim connStrTST As String = "Database=belo_database;Data Source=192.168.100.172;Port=3306;User Id=root ;Password=belo;UseCompression=True;Convert Zero Datetime=True"
    Dim connStrBMG As String = "Database=belo_database;Data Source=192.168.100.250;Port=3306;User Id=root ;Password=webdeveloper;UseCompression=True;Convert Zero Datetime=True"
Private Sub Form2_Load( sender As Object,  e As EventArgs) Handles MyBase.Load
        listview1.Items.Clear()
        listview1.Columns.Clear()
        listview1.Columns.Add("Ref No",100)
        listview1.Columns.Add("Accountno",100)
        listview1.Columns.Add("Accountmain",150)
        listview1.Columns.Add("Memberdate",105)
        listview1.Columns.Add("Fullname",200)
        listview1.Columns.Add("Pointsearned",100)
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
            If txtTop.Text = "" then
                xTop = ""
            End If
            Dim cmdSelect As New OleDbCommand("SELECT [Rec No],Accountno, Accountmain,Memberdate,Firstname, Lastname,Pointsearned FROM " & fi.Name, conn)
            Dim cmdSelect1 As New OleDbCommand("SELECT TOP " & xTop & " [Rec No],Accountno,Accountmain, Memberdate,Firstname, Lastname,Pointsearned FROM " & fi.Name &" order by Pointsearned desc", conn)
            'Dim xTopRange As String() = xTop.Split("-")
            'Dim cmdSelect2 As New OleDbCommand("SELECT TOP " & xTopRange(1).ToString & " [Rec No],Accountno,Accountmain, Memberdate,Firstname, Lastname,Pointsearned FROM " & fi.Name &" order by Pointsearned desc " _
            '                                   & " EXCEPT SELECT TOP " & xTopRange(0).ToString & " [Rec No],Accountno,Accountmain, Memberdate,Firstname, Lastname,Pointsearned FROM " & fi.Name &" order by Pointsearned desc ",conn)

            'Dim readers As OleDbDataReader 
            Dim adapter1 As New OleDbDataAdapter
            If Len(xTop) > 0 then
                adapter1.SelectCommand = cmdSelect1
            ElseIf neg = True then
                Dim cmdSelectNeg As New OleDbCommand("SELECT [Rec No],Accountno, Accountmain,Memberdate,Firstname, Lastname,Pointsearned from " & fi.Name & " where Pointsearned < 0 order by pointsearned asc",conn )
                adapter1.SelectCommand = cmdSelectNeg
            ElseIf negDB =True then
                Dim cmdSelectNegDB As New OleDbCommand("SELECT [Rec No],Accountno, Accountmain,Memberdate,Firstname, Lastname,Pointsearned from " & fi.Name & " where Pointsearned < 0 order by pointsearned asc",conn )
                adapter1.SelectCommand = cmdSelectNegDB
            Elseif xTop = ""
                adapter1.SelectCommand = cmdSelect
            End If

            Dim ds As New DataSet
            Try
                adapter1.Fill(ds, "DATA")
                DataGridView1.DataSource = ds.Tables(0)
                DataGridView1.Columns(0).Width = 100
                DataGridView1.Columns(1).Width = 100
                DataGridView1.Columns(2).Width = 120
                DataGridView1.Columns(3).Width = 100
                DataGridView1.Columns(4).Width = 100
                DataGridView1.Columns(5).Width = 100
                DataGridView1.Columns(6).Width = 100
                ProgBar.Maximum = DataGridview1.RowCount
                StatCount.Text = "Count : " & DataGridview1.RowCount - 1 
                'cmdCheck_Click(Me,EventArgs.Empty)
                conn.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
                GoTo Dist
            End Try
            
        Else
            MsgBox("Nothing to Display!",vbExclamation + vbOKOnly,"No Records found!")
        End If 
End Sub

Private Sub cmdCheck_Click( sender As Object,  e As EventArgs) Handles cmdCheck.Click
        ProgBar.Value = 0
        Dim xCard As String
        Dim xPoints As Integer
        Dim Total As Integer = DataGridView1.RowCount 
        txtAccountno.Text = "" : txtMemberDate.Text = "" : xRec = "" : fname = "" : lname = ""
        Try 
            For i As Integer = 0 to Total
            ProgBar.Value = ProgBar.Value + 1
            StatProgress.Text = "Progress : " & i
            'ProgBar.MarqueeAnimationSpeed = 50
            'Label2.Text = "Process: " & ProgBar.Value & "/" & Total
            
                If negDB =True then
                    xCard = DataGridView1.Rows(i).Cells("cardno").Value.ToString
                    If Mid(xCard,1,2) = "WV" then
                        xCard = Replace(xCard,"WV","",1) 
                    End If
                    xPoints = DataGridView1.Rows(i).Cells("total_points").Value
                    ReadInverseChecker(xCard,xPoints,i)
                else
                    xCard = DataGridView1.Rows(i).Cells("Accountno").Value.ToString
                    If Mid(xCard,1,2) = "WV" then
                        xCard = Replace(xCard,"WV","",1) 
                    End If
                    xPoints = DataGridView1.Rows(i).Cells("Pointsearned").Value
                    ReadChecker(xCard,xPoints,i)
                End If
            
            Next
        Catch ex As Exception
            If InStr(1,ex.Message,"Object reference not set to an instance of an object.") then
                
            else
                MsgBox(ex.Message)
            End If
        End Try
        MsgBox("Done")
End Sub

Private Sub ReadChecker(ByVal CardNum As string, ByVal TotalPoints As Integer, ByVal steps As Integer ) 
        'Dim sr As New StreamReader(Application.StartupPath & "\SqlConnection.dll")
recon:
        Dim query As String = "select * from `Loyalty_Card_Info` where CardNo like '%" & CardNum & "%' or patientid like '%"& CardNum &"%' and card_status not in ('not issued','notissued')"
        'Dim connStrSql As String = sr.ReadLine
        Dim connection As New MySqlConnection(connStrBMG)
        'Dim connection1 As New MySqlConnection(connStrBMG)
        Dim cmd As new MySqlCommand(query,connection)
        Dim reader As MySqlDataReader
        Try
            connection.Open
            reader = cmd.ExecuteReader
            If reader.HasRows then
                While reader.Read
                    If TotalPoints <> Val(reader.Item("total_points").ToString) then
                        If txtAccountno.Text = "" then
                            xRec = DataGridView1.Rows(steps).Cells("Rec No").Value.ToString & "|"
                            txtAccountno.Text = DataGridView1.Rows(steps).Cells("AccountNo").Value.ToString & "|"
                            AccMain = DataGridView1.Rows(steps).Cells("Accountmain").Value.ToString & "|"
                            dt = DataGridView1.Rows(steps).Cells("Memberdate").Value.ToString 
                            txtMemberDate.Text = Convert.ToDateTime(dt) & "|"
                            fname = DataGridView1.Rows(steps).Cells("Firstname").Value.ToString & "|"
                            lname = DataGridView1.Rows(steps).Cells("Lastname").Value.ToString & "|"
                            xTotal = DataGridView1.Rows(steps).Cells("Pointsearned").Value.ToString & "|"
                            Points = reader.Item("total_points").ToString & "|"
                            CountSteps = CountSteps + 1
                        Else
                            xRec = xRec & DataGridView1.Rows(steps).Cells("Rec No").Value.ToString & "|"
                            txtAccountno.Text = txtAccountno.Text & DataGridView1.Rows(steps).Cells("AccountNo").Value.ToString & "|"
                            AccMain = AccMain & DataGridView1.Rows(steps).Cells("Accountmain").Value.ToString & "|"
                            dt = DataGridView1.Rows(steps).Cells("Memberdate").Value.ToString 
                            txtMemberDate.Text = txtMemberDate.Text & Convert.ToDateTime(dt) & "|"
                                if len(DataGridView1.Rows(steps).Cells("Firstname").Value.ToString) > 0
                                fname = fname & DataGridView1.Rows(steps).Cells("Firstname").Value.ToString & "|"
                                Else
                                fname = fname & "NONE|"
                                end if
                            lname = lname & DataGridView1.Rows(steps).Cells("Lastname").Value.ToString & "|"
                            xTotal = xTotal & DataGridView1.Rows(steps).Cells("Pointsearned").Value.ToString & "|"
                            Points = Points & reader.Item("total_points").ToString & "|"
                            CountSteps = CountSteps + 1
                        End If
                    else
                        'Ignore
                    End If
                End While
            Else
                MsgBox("Meow")
            End If
        Catch ex As Exception
            Dim err As String
            err = ex.Message
            
            If InStr(1,ex.Message,"Object reference not set to an instance of an object.") then
            elseIf InStr(1,ex.Message,"Reading from the stream has failed") then
                MsgBox("Connection has been lost. Reconnecting...")
                GoTo recon
            else
                MsgBox(ex.Message)
            End If
        End Try
        connection.Close
End Sub

Private Sub cmdExcel_Click( sender As Object,  e As EventArgs) Handles cmdExcel.Click
        Dim stepx As Integer = Val(Label2.Text)
        Dim xRecord As String() = xRec.Split("|")
        Dim xAccountMain As String() = AccMain.Split("|")
        Dim xGtotal As String() = xTotal.Split("|") : Dim xPoints As String() = Points.Split("|")
        Dim xAccountNo As String() = txtAccountno.Text.Split("|") 
        Dim xMemberDate As String() = txtMemberDate.Text.Split("|")
        Dim xFirstname As String() = fname.Split("|") : Dim xLastname As String() = lname.Split("|")
        Dim xFullname As String() = fullname.Split("|") 
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

        sheet.Cells(1,1) = "UNMATCH RECORDS FOUND AS OF " & Now().ToString("MM/dd/yyyy") : sheet.Cells(1,1).Font.Name = "Calibri"
        sheet.Cells(1,1).Font.Bold = True : sheet.Cells(1,1).Font.Size = 22
        sheet.Cells(1,1).Font.color = Color.Red
        sheet.Range("A1:G1").MergeCells = True

        sheet.Cells(2,1) = "Record No"      : sheet.Cells(2,1).Font.Bold = True
        sheet.Cells(2,2) = "Card Number"    : sheet.Cells(2,2).Font.Bold = True
        sheet.Cells(2,3) = "Patient ID"     : sheet.Cells(2,3).Font.Bold = True : sheet.Cells(2,3).Alignment = "Left"
        sheet.Cells(2,4) = "Member Date"    : sheet.Cells(2,4).Font.Bold = True
        sheet.Cells(2,5) = "Fullname"       : sheet.Cells(2,5).Font.Bold = True
        sheet.Cells(2,6) = "Points in POS"  : sheet.Cells(2,6).Font.Bold = True
        sheet.Cells(2,7) = "Points in DB"   : sheet.Cells(2,7).Font.Bold = True
        '
        '
        For i As Integer = 0 to CountSteps
            stepx = i + 3
        sheet.Cells(stepx,1) = xRecord(i)
        sheet.Cells(stepx,2) = xAccountNo(i)
        sheet.Cells(stepx,3) = xAccountMain(i)
        sheet.Cells(stepx,4) = xMemberDate(i)
        If i > -1 and Len(fullname) > 1 then
        sheet.Cells(stepx,5) = xFullname(i) 
        Else
        sheet.Cells(stepx,5) = xFirstname(i) & " " & xLastname(i)
        End If
        sheet.Cells(stepx,6) = xGtotal(i)
        sheet.Cells(stepx,7) = xPoints(i)
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

Private Sub cmdTop100_Click( sender As Object,  e As EventArgs) Handles cmdTop100.Click
        xTop = "100"
        cmdDisplay_Click(Me,EventArgs.Empty)
End Sub

Private Sub cmdTop1000_Click( sender As Object,  e As EventArgs) Handles cmdTop1000.Click
        xTop = "1000"
        cmdDisplay_Click(Me,EventArgs.Empty)
End Sub

Private Sub cmdTop10000_Click( sender As Object,  e As EventArgs) Handles cmdTop10000.Click
        xTop = "10000"
        cmdDisplay_Click(Me,EventArgs.Empty)
End Sub

    Private Sub txtTop_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtTop.KeyPress
    '97 - 122 = Ascii codes for simple letters
    '65 - 90  = Ascii codes for capital letters
    '45 = Ascii code for negative/hyphen
    '48 - 57  = Ascii codes for numbers
    '8 = Ascii code for backspace

    If Asc(e.KeyChar) <> 8 Then
        'If Asc(e.KeyChar) = 45 then
        '    If InStr(1,txtTop.Text,"-") then
        '        e.Handled = True
        '    Else
        '        e.Handled = False
        '    End If
        'Else
            If  Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
            e.Handled = True
            End If 
        'End If
    End If
    End Sub

    Private Sub txtTop_KeyUp(sender As Object, e As KeyEventArgs) Handles txtTop.KeyUp
        If e.KeyCode = Keys.Enter then
            If txtTop.Text <> "" then
                xTop = txtTop.Text
                cmdDisplay_Click(Me,EventArgs.Empty)
            End If
        End If
    End Sub

Private Sub cmdNegPOS_Click( sender As Object,  e As EventArgs) Handles cmdNegPOS.Click
        neg = True
        cmdDisplay_Click(Me,EventArgs.Empty)
End Sub

Private Sub cmdNegDB_Click( sender As Object,  e As EventArgs) Handles cmdNegDB.Click
reconn:
        negDB = True
        strSql = "SELECT id,cardno,patientid,activation_date,(SELECT CONCAT(`firstname` , ' ' , `lastname`) FROM `patient_info` " _
        & " WHERE `patient_info`.`PatientID` = `loyalty_card_info`.`PatientID`) AS Fullname, total_points " _
        & " FROM `Loyalty_Card_Info` where total_points < 0 and card_status NOT IN ('not issued','notissued')"
        Dim connection As New MySqlConnection(connStrBMG)
        Dim cmd As New MySqlCommand(strSql,connection)
        Dim adapter As New MySqlDataAdapter(cmd)
        Dim ds As New DataSet
        Try
            connection.Open
            adapter.Fill(ds,"Loyalty_Card_Info")
            DataGridview1.DataSource = ds.Tables(0)
            StatCount.Text = "Count : " & DataGridView1.RowCount - 1
            DataGridView1.Columns(0).HeaderText = "Ref No"
            DataGridView1.Columns(1).HeaderText = "Card Number"
            DataGridView1.Columns(2).HeaderText = "Patient ID"
            DataGridView1.Columns(3).HeaderText = "Member Date"
            DataGridView1.Columns(4).HeaderText = "Fullname"
            DataGridView1.Columns(5).HeaderText = "Points Earned"
        Catch ex As Exception
            If InStr(1,ex.Message,"Unable to connect to any of the specified MySQL") then
                MsgBox("Connection has been lost. Reconnecting...",vbOKOnly + vbCritical,"Reconnecting..")
                GoTo reconn
            Elseif InStr(1,ex.Message,"Reading from the stream has failed") then
                MsgBox("Connection has been lost. Reconnecting...",vbOKOnly + vbCritical,"Reconnecting..")
                GoTo reconn
            Else
                MsgBox(ex.Message)
            End If
        End Try
End Sub

Private Sub ReadInverseChecker(ByVal CardNum As string, ByVal TotalPoints As Integer, ByVal steps As Integer ) 
         If txtFileDirectory.Text <> "" then
reconn:     
            Dim FilePath As String = txtFileDirectory.Text 
            Dim fi As New FileInfo(FilePath)
            Dim connectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=Text;Data Source=" & fi.DirectoryName
            Dim conn As New OleDbConnection(connectionString)
            Dim cmd As New OleDbCommand("SELECT [Rec No],Accountno, Accountmain,Memberdate,Firstname, Lastname,Pointsearned FROM " & fi.Name & " where [Accountno] like '" & CardNum & "' or Accountmain like '" & CardNum & "'", conn)
            Dim cmdSelect1 As New OleDbCommand("SELECT TOP " & xTop & " [Rec No],Accountno,Accountmain, Memberdate,Firstname, Lastname,Pointsearned " _
                                               & " FROM " & fi.Name &" where Accountno like '" & CardNum & "' or Accountmain like '" & CardNum & "'", conn)
            Dim reader As OleDbDataReader 
           
            Try
            conn.Open
            reader = cmd.ExecuteReader
            If reader.HasRows then
                While reader.Read
                    'If TotalPoints <> Val(reader("Pointsearned").ToString) then
                        If txtAccountno.Text = "" then
                            xRec = DataGridView1.Rows(steps).Cells("id").Value.ToString & "|"
                            txtAccountno.Text = DataGridView1.Rows(steps).Cells("cardno").Value.ToString & "|"
                            AccMain = DataGridView1.Rows(steps).Cells("patientid").Value.ToString & "|"
                            dt = DataGridView1.Rows(steps).Cells("activation_date").Value.ToString 
                            txtMemberDate.Text = Convert.ToDateTime(dt) & "|"
                            fullname = fullname & DataGridView1.Rows(steps).Cells("Fullname").Value.ToString & "|"
                            xTotal = reader.Item("Pointsearned").ToString & "|"
                            Points = DataGridView1.Rows(steps).Cells("total_points").Value.ToString & "|"
                            CountSteps = CountSteps + 1
                        Else
                            xRec = xRec & DataGridView1.Rows(steps).Cells("id").Value.ToString & "|"
                            txtAccountno.Text = txtAccountno.Text & DataGridView1.Rows(steps).Cells("cardno").Value.ToString & "|"
                            AccMain = AccMain & DataGridView1.Rows(steps).Cells("patientid").Value.ToString & "|"
                            dt = DataGridView1.Rows(steps).Cells("activation_date").Value.ToString 
                            txtMemberDate.Text = txtMemberDate.Text & Convert.ToDateTime(dt) & "|"
                            fullname = fullname & DataGridView1.Rows(steps).Cells("Fullname").Value.ToString & "|"
                            xTotal = xTotal & reader.Item("Pointsearned").ToString & "|"
                            Points = Points & DataGridView1.Rows(steps).Cells("total_points").Value.ToString & "|"
                            CountSteps = CountSteps + 1
                        End If
                    'Else
                        'Ignore
                    'End If
                End While
            End If
            Catch ex As Exception
                Dim err As String
                err = ex.Message
            
                If InStr(1,ex.Message,"Object reference not set to an instance of an object.") then
                elseIf InStr(1,ex.Message,"Reading from the stream has failed") then
                    MsgBox("Connection has been lost. Reconnecting...")
                    GoTo reconn
                else
                    MsgBox(ex.Message)
                End If
            End Try
            conn.Close
        End If
End Sub

Private Sub txtTop_TextChanged( sender As Object,  e As EventArgs) Handles txtTop.TextChanged

End Sub
End Class