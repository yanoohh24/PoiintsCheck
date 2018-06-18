
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports System.Reflection
Imports Mysql.data.mysqlclient
Imports System.IO
Imports System.IO.Ports
Imports System.IO.Compression
Imports System.ComponentModel
Imports System.Threading
Imports System.ServiceProcess

Module Module1
    Public connStrSMS as String
    Public connStrBMG as String
    public Database As string
    Public cn As new MySqlConnection
    Public command As new MySqlCommand
    public adapt as new MySqlDataAdapter
    public Reader as MySqlDataReader
    Public rb as new OleDbConnection
    public rs as new OleDbCommand
    public rst as New OleDbDataAdapter
    Public strSQL as String
    public host As String
    public Username As String
    Public Password As String
 

    sub main()
        host = "127.0.0.1"
        Username = "root"
        Password = ""
        'frmSMS.Show()
        'connStrSMS = "Database=Messages;Data Source=" & Host & ";User Id=" & UserName & ";Password=" & Password & ";UseCompression=True;Connection Timeout=28800"
        'connStrBMG = "Database=belo_database;Data Source=" & Host & ";User Id=" & UserName & ";Password=" & Password & ";UseCompression=True;Connection Timeout=28800;Convert Zero Datetime=True"
        
    End Sub

End Module
