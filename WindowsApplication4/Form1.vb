
Imports System.Data.OleDb
Imports Mysql.data.mysqlclient
Imports System.IO
Imports System.IO.Ports
Imports System.IO.Compression
Imports System.ComponentModel
Imports System.ComponentModel.Design.Serialization
Imports System.Configuration
Imports System.Threading
Imports Microsoft.Office.Interop
Imports System.Reflection
Imports System.ServiceProcess
Imports System.Management
Imports System.Text.RegularExpressions

Public Class Form1
    dim rcvdata as string = ""

    Public Function ModemsConnected() As String
        Dim modems As String = ""
        Try
            Dim searcher As New ManagementObjectSearcher( _
                "root\CIMV2", _
                "SELECT * FROM Win32_POTSModem")

            For Each queryObj As ManagementObject In searcher.Get()
                If queryObj("Status") = "OK" Then
                    modems = modems & (queryObj("AttachedTo") & " - " & queryObj("Description") & "***")
                End If
            Next
        Catch err As ManagementException
            MessageBox.Show("An error occurred while querying for WMI data: " & err.Message)
            Return ""
        End Try
        Return modems
    End Function

Private Sub Form1_Load( sender As Object,  e As EventArgs) Handles MyBase.Load
    RunningTime.Start()
    lblRunningTime.Text = TimeOfDay()
    'LoadInbox()
    GroupBox1.show()
    
    'Load COM (Portal for Sending SMS)
    Try
        dim Ports as String() = SerialPort.GetPortNames
        dim Port as String
        for each port in ports
            cmbCOM.Items.Add(port)
        Next Port
        cmbCOM.SelectedIndex = 0
    Catch ex As Exception
        MsgBox("NO AVAILABLE PORT TO SELECT! PLEASE CHECK PORT CONNECTION",vbOKOnly + vbCritical, "MISSSING PORT!")

    End Try

    Try
        dim forts() as String
        forts = split(ModemsConnected(),"***")
        for i as Integer = 0 to forts.Length = 2
            ComboBox1.Items.Add(forts(i))
        Next
    Catch ex As Exception

    End Try
End Sub

Public Class SMSCOMMS
    Private WithEvents SMSPort As SerialPort
    Private SMSThread As Thread
    Private ReadThread As Thread
    Shared _Continue As Boolean = False
    Shared _ContSMS As Boolean = False
    Private _Wait As Boolean = False
    Shared _ReadPort As Boolean = False
    Public Event Sending(ByVal Done As Boolean)
    Public Event DataReceived(ByVal Message As String)

    Public sub New(byRef COMMPORT As String)
        'initialize all values
        SMSPort = New SerialPort
        With SMSPort
            .PortName = COMMPORT
            .BaudRate = 19200
            .Parity = Parity.None
            .DataBits = 8
            .StopBits = StopBits.One
            .Handshake = Handshake.RequestToSend
            .DtrEnable = True
            .RtsEnable = True
            .NewLine = vbCrLf
        End With
        
    End sub
 

    Public Sub Open()
        If Not (SMSPort.IsOpen = True) Then
            SMSPort.Open()
        End If
    End Sub

    Public Sub Close()
        If SMSPort.IsOpen = True Then
            SMSPort.Close()
        End If
    End Sub
End Class


Private Sub RunningTime_Tick( sender As Object,  e As EventArgs) Handles RunningTime.Tick
        lblRunningTime.Text = TimeOfDay.ToString("hh:mm:ss tt")
End Sub

Public sub LoadInbox
    ListView1.View = view.Details
    ListView1.Clear()
    ListView1.Columns.Add("NO",40, HorizontalAlignment.Left)
    ListView1.Columns.Add("ID#",55, HorizontalAlignment.Left)
    ListView1.Columns.Add("SMS",350, HorizontalAlignment.Left)
    ListView1.Columns.Add("MOBILE",120, HorizontalAlignment.Left)
    
End Sub

    public sub LoadMessage
        strsql = "Select * from messages"
        dim cn as New  MySqlConnection(strSQL)
        dim command as new MySqlCommand(strsql,cn)
        dim Reader as MySqlDataReader
        dim iGLCount as long = 0
        cn.Open()
        Reader = command.ExecuteReader()
        ListView1.Items.Clear()
        ' if rs.eof
        if Reader.HasRows = true then
            while reader.Read
                iGLCount += 1
                Dim ls As New ListViewItem(iGLCount.ToString())
                ls.SubItems.Add(reader.Item("id").ToString())
                ls.SubItems.Add(reader.Item("body").ToString())
                ls.SubItems.Add(reader.Item("sender").ToString())
                ListView1.Items.Add(ls)
            End While
        End If
    End Sub

Private Sub SettingToolStripMenuItem_Click( sender As Object,  e As EventArgs) Handles SettingToolStripMenuItem.Click
        GroupBox1.Show()
End Sub
Private Sub Button1_Click( sender As Object,  e As EventArgs) Handles Button1.Click
        '== CONNECTION ==
        Try
            With SerialPort1
                .PortName = cmbcom.Text
                .BaudRate = 19200
                .Parity = IO.Ports.Parity.None
                .DataBits = 8
                .StopBits = IO.Ports.StopBits.One
                .Handshake = IO.Ports.Handshake.RequestToSend
                .ReceivedBytesThreshold = 1
                .NewLine = vbCr
                .DtrEnable = True
                .RtsEnable = True
                .ReadTimeout = 1000
                .Open()

                If .IsOpen() Then
                    lblPort.Text = "Port : "
                    lblPort.Text = lblPort.Text + cmbCOM.Text
                    lblPort.BackColor = color.Green
                    GroupBox1.Hide()
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message,vbOKOnly + vbInformation,"Error!")
            me.Dispose()
        End Try
    
End Sub

Private Sub cmdSend_Click( sender As Object,  e As EventArgs) Handles cmdSend.Click
        Try
            If SerialPort1.IsOpen
            with SerialPort1
                .WriteLine("AT" & vbCrLf)
                .WriteLine("AT+CMGF=1" & vbCrLf) 'set command message format to text mode(1)
                .WriteLine("AT+CMGS= " & Chr(34) & txtMobile.Text & Chr(34) & vbCrLf) 'textbox is the destination no
                .WriteLine(RichTextBox1.Text & Chr(26))
                Threading.Thread.Sleep(1000)
                MsgBox(rcvdata.ToString)
            End With
                if rcvdata.tostring.Contains(">")
                    MsgBox("Message Sent",,"")
                else
                    MsgBox("Something Went Wrong",,"")
                End If
            End If  
        Catch ex As Exception
            'IGNORED
        End Try
End Sub
    
Private Sub ToolStrip1_ItemClicked( sender As Object,  e As ToolStripItemClickedEventArgs) Handles ToolStrip1.ItemClicked

End Sub

Private Sub serialport_datareceived(ByVal sender As object, ByVal  e As system.IO.Ports.SerialDataReceivedEventArgs) handles SerialPort1.DataReceived
 '      == SERIAL PORT ==       
        dim datain as String = ""
        dim numbytes as Integer = SerialPort1.BytesToRead
        for i as Integer = 1 to numbytes
            datain & = Chr(SerialPort1.ReadChar)
        Next
        test(datain)
End Sub

Private Sub test (ByVal indata As String)
    rcvdata & = indata
End Sub

Private Sub cmdRead_Click( sender As Object,  e As EventArgs) Handles cmdRead.Click
        try 
            with SerialPort1
                rcvdata = ""
                .Write("AT" & vbCrLf)
                Threading.Thread.Sleep(1000)
                .Write("AT+CMGF=1" & vbCrLf)
                Threading.Thread.Sleep(1000)
                .Write("AT+CPMS""SM""" & vbCrLf)
                Threading.Thread.Sleep(1000)
                .Write("AT+CMGL=""ALL""" & vbCrLf)
                Threading.Thread.Sleep(1000)
                'MsgBox(rcvdata.ToString)
                ReadMessage()
            End With
                
                
        Catch ex As Exception

        End Try
End Sub

public sub ReadMessage()
        dim lineoftext as String
        dim i as Integer
        dim arraytextfile() as String
        Lineoftext = rcvdata.ToString
        arraytextfile = Split(lineoftext,"+CMGL",,CompareMethod.Text)
        for i = 0 To UBound(arraytextfile)
            'MsgBox(arraytextfile(i))
            Dim input as String = arraytextfile(i)
            dim result() as String
            dim pattern as String = "(:)|(,"")|("","")"
            result = Regex.Split(input,pattern)
            dim lvi as new listview
            dim concat() as string
            with ListView1.items.add("null")
                'Index
                .subitems.AddRange(New String() {result(2)})
                'For Status
                .subitems.AddRange(New String() {result(4)})
                'For Status
                '.subitems.AddRange(New String(){result(6)})
                dim my_string, position as String
                my_string = result(6)
                position = my_string.Length - 2
                my_string = my_string.Remove(position,2)
                .SubItems.Add(my_string)
                'date and time
                concat = new String() {result(9) & result(10) & result(11) & result(12)}
                .SubItems.Addrange(concat)
                'For Messages
                dim lineoftexts as String
                dim arrayoftexts() as String
                lineoftexts = arraytextfile(i)
                arrayoftexts = Split(lineoftexts,"+32",,CompareMethod.Text)
                .SubItems.Add(arrayoftexts(i))
            End With
        Next
        End Sub

Private Sub ToolStripStatusLabel3_Click( sender As Object,  e As EventArgs) Handles ToolStripStatusLabel3.Click

End Sub

Private Sub Button2_Click( sender As Object,  e As EventArgs) 
        
End Sub

Private Sub GroupBox1_Enter( sender As Object,  e As EventArgs) Handles GroupBox1.Enter

End Sub

'Private Function ModemsConnected() As String 
'Throw New NotImplementedException 
' End Function 

Private Sub ExitToolStripMenuItem_Click( sender As Object,  e As EventArgs) Handles ExitToolStripMenuItem.Click
        me.Dispose()
End Sub

Private Sub Button2_Click_1( sender As Object,  e As EventArgs) 

End Sub
End Class 'What Class? 




