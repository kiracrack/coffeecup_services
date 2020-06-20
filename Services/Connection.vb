Imports MySql.Data.MySqlClient ' this is to import MySQL.NET
Imports System.Data
Imports System.Management
Imports Microsoft.VisualBasic
Imports System.Net.Mail
Imports System.Text
Imports System.IO
Imports Microsoft.Win32
Imports System.Security.Cryptography
Imports System.Security.Principal
Imports System.Windows.Forms

Module Connection
    Public conn As New MySqlConnection 'for MySQLDatabase Connection
    Public msda As MySqlDataAdapter 'is use to update the dataset and datasource
    Public dst As New DataSet 'miniature of your table - cache table to client
    Public e_msda As MySqlDataAdapter 'is use to update the dataset and datasource
    Public e_dst As New DataSet 'miniature of your table - cache table to client

    Public rpt_msda As MySqlDataAdapter 'is use to update the dataset and datasource
    Public rpt_dst As New DataSet 'miniature of your table - cache table to client




    Public com As New MySqlCommand
    Public rst As MySqlDataReader
    Public cb As MySqlCommandBuilder

    ' LOCALHOST
    Public sqlservername As String
    Public sqlserver As String
    Public sqlport As String
    Public sqluser As String
    Public sqlpass As String
    Public sqldatabase As String
    Public sqljoinbase As String
    Public conString As String
    Public file_conn As String = Application.StartupPath.ToString & "\Coffeecup.conn"
    Public ConnectedServer As Boolean = False
     
    Public Function ConnectServer() As Boolean
        Dim strSetup As String = ""
        Dim sr As StreamReader = File.OpenText(file_conn)
        Dim br As String = sr.ReadLine() : sr.Close()
        strSetup = DecryptTripleDES(br) : Dim cnt As Integer = 0
        For Each word In strSetup.Split(New Char() {","c})
            If cnt = 0 Then
                sqlservername = word
            ElseIf cnt = 1 Then
                sqlserver = word
            ElseIf cnt = 2 Then
                sqlport = word
            ElseIf cnt = 3 Then
                sqluser = word
            ElseIf cnt = 4 Then
                sqlpass = word
            ElseIf cnt = 5 Then
                sqldatabase = word
            ElseIf cnt = 6 Then
                sqljoinbase = word
            End If
            cnt = cnt + 1
        Next
        Try
            conn = New MySql.Data.MySqlClient.MySqlConnection
            conn.ConnectionString = "server=" & sqlserver & "; Port=" & sqlport & "; user id=" & sqluser & " ; password=" & sqlpass & " ; database=" & sqldatabase & " ; Allow Zero Datetime=True ; Connection Timeout=28800 ; allow user variables=true"
            com.CommandTimeout = 28800
            com.Connection = conn
            conn.Open()
            RecordLog("MySQL server " & sqlserver & " successfully Connected")
            LoadGeneralSettings()
        Catch errMYSQL As MySqlException
            CheckConnection()
        End Try
        Return True
    End Function
    Public Sub CheckConnection()
        If ConnectServer() = True Then
            Dim ServicesToRun() As System.ServiceProcess.ServiceBase
            ServicesToRun = New System.ServiceProcess.ServiceBase() {New CoffeecupServices()}
            System.ServiceProcess.ServiceBase.Run(ServicesToRun)
        Else
            CheckConnection()
        End If
    End Sub
End Module
