Imports System.Windows.Forms
Imports System.IO
Imports System.ServiceProcess
Imports System.ComponentModel
Imports System.Security.Permissions

<PermissionSet(SecurityAction.Demand, Name:="FullTrust")> _
<System.Runtime.InteropServices.ComVisibleAttribute(True)> _
Public Class CoffeecupServices

    Public timer As System.Timers.Timer = New System.Timers.Timer()
    Public bwWeekly As BackgroundWorker = New BackgroundWorker
    Public bwMonthly As BackgroundWorker = New BackgroundWorker
    Protected Overrides Sub OnStart(ByVal args() As String)
        If ConnectServer() = True Then
            timer.Interval = 30000 ' 1 second
            AddHandler timer.Elapsed, AddressOf Me.OnTimer
            timer.Enabled = True
            timer.Start()
            RecordLog("Coffeecup mail service started")

            bwWeekly.WorkerSupportsCancellation = True
            AddHandler bwWeekly.DoWork, AddressOf bwWeekly_DoWork

            bwMonthly.WorkerSupportsCancellation = True
            AddHandler bwMonthly.DoWork, AddressOf bwMonthly_DoWork

        Else
            CheckConnection()
        End If
    End Sub

    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.
        RecordLog("Coffeecup mail service Stop")
    End Sub
    Private Sub OnTimer(sender As Object, e As Timers.ElapsedEventArgs)
        Try
            If DecryptTripleDES(GlobalEngineCode) = getMacAddress() Then
                If CDate(Now.ToShortDateString) > CDate(DecryptTripleDES(GlobalSystemDate)) Then
                    com.CommandText = "update tblsystemlicense set systemdate='" & EncryptTripleDES(ConvertDate(CDate(Now.ToShortDateString))) & "'" : com.ExecuteNonQuery()
                    GlobalSystemDate = EncryptTripleDES(ConvertDate(CDate(Now.ToShortDateString)))
                    RecordLog("System date updated to " & ConvertDate(CDate(Now.ToShortDateString)))
                End If
            End If
            If globalEmailNotification = True Then
                EmailAllPendingNotification()
            End If

            If Not bwWeekly.IsBusy And Not bwMonthly.IsBusy Then
                If globalEmailNotification = True Then
                    If GlobalEmailNotifyMonthlySummary = True Then
                        If countqry("tblaccounts", "notifymonthlysummary=1") > 0 Then
                            If Globalweeklyreportdate <> "" Then
                                Globalweeklyreportdate = qrysingledata("weeklyreportdate", "weeklyreportdate", "tblgeneralsettings")
                                If CDate(Now.ToShortDateString) >= CDate(Globalweeklyreportdate) Then
                                    bwWeekly.RunWorkerAsync()
                                End If
                            End If
                        End If
                    End If
                End If
            End If

            If Not bwWeekly.IsBusy And Not bwMonthly.IsBusy Then
                If globalEmailNotification = True Then
                    If GlobalEmailNotifyMonthlySummary = True Then
                        If countqry("tblaccounts", "notifymonthlysummary=1") > 0 Then
                            If Globalmonthlyreportdate <> "" Then
                                Globalmonthlyreportdate = qrysingledata("monthlyreportdate", "monthlyreportdate", "tblgeneralsettings")
                                If CDate(Now.ToShortDateString) >= CDate(Globalmonthlyreportdate) Then
                                    bwMonthly.RunWorkerAsync()
                                End If
                            End If
                        End If
                    End If
                End If
            End If


        Catch ex As Exception
            RecordLog("Error OnTimer: " + ex.Message)
            CheckConnection()
        Finally
            timer.Start()
        End Try
    End Sub


#Region "Live refresh contact list - background worker"
    Private Sub bwWeekly_DoWork(ByVal sender As Object, ByVal e As DoWorkEventArgs)
        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
        For i = 1 To 1
            If bwMonthly.CancellationPending = True Then
                e.Cancel = True
                Exit For
            End If
            SendEmailReport("Weekly", Globalweeklyreportdate)
        Next
    End Sub

    Private Sub bwMonthly_DoWork(ByVal sender As Object, ByVal e As DoWorkEventArgs)
        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
        For i = 1 To 1
            If bwWeekly.CancellationPending = True Then
                e.Cancel = True
                Exit For
            End If
            SendEmailReport("Monthly", Globalmonthlyreportdate)
        Next
    End Sub
#End Region

    Protected Overrides Sub OnContinue()
        Application.Restart()
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Public Sub New()
        InitializeComponent()
    End Sub

End Class
