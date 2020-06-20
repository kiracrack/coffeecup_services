Imports MySql.Data.MySqlClient
Imports System.Windows.Forms

Module General_Settings
    'Declaration of email settings
    Public globalEmailNotification As Boolean
    Public GlobalEmailNotifyMonthlySummary As Boolean
    Public Globalweeklyreportdate As String
    Public Globalmonthlyreportdate As String
    Public GlobalOrganizationLogoURL As String
    Public GlobalCompanyName As String
    Public GlobalEngineCode As String
    Public GlobalSystemDate As String
    Public globalsmtpHost As String
    Public globalsmtpPort As String
    Public globalsslEnable As Boolean
    Public globalserverEmailAddress As String
    Public globaltargetEmailAddress As String
    Public globalemailPassword As String

    Public Sub LoadGeneralSettings()
        Try
            com.CommandText = "select * from tblgeneralsettings"
            rst = com.ExecuteReader
            While rst.Read
                globalEmailNotification = rst("enableemailnotification")
                globalsmtpHost = rst("smtphost").ToString()
                globalsmtpPort = rst("smptport").ToString()
                globalsslEnable = rst("smtpsslenable")
                globalserverEmailAddress = rst("serveremailaddress").ToString()
                globalemailPassword = DecryptTripleDES(rst("serverpassword").ToString())
                GlobalEmailNotifyMonthlySummary = rst("emailnotifymonthlysummary")
                Globalweeklyreportdate = rst("weeklyreportdate").ToString
                Globalmonthlyreportdate = rst("monthlyreportdate").ToString
            End While
            rst.Close()

            com.CommandText = "select * from tblcompanysettings where defaultcompany=1"
            rst = com.ExecuteReader
            While rst.Read
                GlobalCompanyName = rst("companyname").ToString
                GlobalOrganizationLogoURL = rst("logourl").ToString
            End While
            rst.Close()

            com.CommandText = "select * from tblsystemlicense"
            rst = com.ExecuteReader
            While rst.Read
                GlobalEngineCode = rst("enginecode").ToString
                GlobalSystemDate = rst("systemdate").ToString
            End While
            rst.Close()
            If GlobalSystemDate = "" Then
                com.CommandText = "update tblsystemlicense set systemdate='" & EncryptTripleDES(ConvertDate(CDate(Now.ToShortDateString))) & "'" : com.ExecuteNonQuery()
            End If
        Catch ex As Exception
            RecordLog("Error Notification: " + ex.Message)
        End Try
    End Sub
    Public Function countqry(ByVal tbl As String, ByVal where As String)
        Dim cnt As Integer = 0
        com.CommandText = "select count(*) as cnt from " & tbl & " where " & where
        rst = com.ExecuteReader
        While rst.Read
            cnt = rst("cnt")
        End While
        rst.Close()
        Return cnt
    End Function

    Public Function countrecord(ByVal tbl As String)
        Dim cnt As Integer = 0
        com.CommandText = "select count(*) as cnt from " & tbl & " "
        rst = com.ExecuteReader
        While rst.Read
            cnt = rst("cnt")
        End While
        rst.Close()
        Return cnt
    End Function

    Public Function qrysingledata(ByVal field As String, ByVal fqry As String, ByVal tbl As String)
        Dim def As String = ""
        Try
            com.CommandText = "select " & fqry & " from " & tbl : rst = com.ExecuteReader
            While rst.Read
                def = rst(field).ToString
            End While
            rst.Close()
        Catch errMYSQL As MySqlException

        End Try
        Return def
    End Function

     
    Public Function SendEmailReport(ByVal reportType As String, ByVal ReportDate As String)
        Try
            Dim Template As String = Application.StartupPath.ToString & "\Templates\Email\EmailReportTemplate.html"
            If System.IO.File.Exists(Template) = True Then
                Dim SaveLocation As String = ""
                If reportType = "Weekly" Then
                    SaveLocation = Application.StartupPath.ToString & "\Transaction\Email\Weekly\" & ConvertDate(CDate(ReportDate)) & " Weekly.html"
                ElseIf reportType = "Monthly" Then
                    SaveLocation = Application.StartupPath.ToString & "\Transaction\Email\Monthly\" & ConvertDate(CDate(ReportDate)) & " Monthly.html"
                End If

                If ReportDate.Length > 0 Then
                    Dim EmailList As String = qrysingledata("emails", "group_concat(emailaddress) as emails", "tblaccounts where notifymonthlysummary=1")
                    If EmailList.Length > 1 Then
                        Dim ReportTitle As String = "" : Dim ReportDetails As String = "" : Dim HtmlBody As String = ""

                        If System.IO.File.Exists(SaveLocation) = True Then
                            System.IO.File.Delete(SaveLocation)
                        End If
                        My.Computer.FileSystem.CopyFile(Template, SaveLocation)

                        HtmlBody = ""
                        ReportTitle = UCase(reportType) & " REPORT SUMMARY"

                        If reportType = "Weekly" Then
                            ReportDetails = "Date From: " & CDate(ReportDate).AddDays(-7) & "<br/>" _
                                        + " Date To: " & CDate(ReportDate).AddDays(-1)
                        ElseIf reportType = "Monthly" Then
                            ReportDetails = "Report Month: " & CDate(ReportDate).AddMonths(-1).ToString("MMMM yyyy")
                        End If

                        rpt_msda = Nothing : rpt_dst = New DataSet
                        rpt_msda = New MySqlDataAdapter("select * from tblemailreportnotification where rptgroup='" & LCase(reportType) & "' and enablenotify=1 order by rptorder asc", conn)
                        rpt_msda.Fill(rpt_dst, 0)
                        For x = 0 To rpt_dst.Tables(0).Rows.Count - 1
                            With (rpt_dst.Tables(0))
                                Dim ReportTable As String = "" : Dim ReportHeader As String = "" : Dim ReportColHeader As String = "" : Dim ReportRowValue As String = "" : Dim tmpField As String = ""
                                com.CommandText = "CALL sp_emailreportnotification('" & .Rows(x)("shortcode").ToString() & "')" : com.ExecuteNonQuery()

                                If countrecord(.Rows(x)("shortcode").ToString()) > 1 Then
                                    Dim cnt As Integer = 0
                                    com.CommandText = "show fields from " & .Rows(x)("shortcode").ToString() & "" : rst = com.ExecuteReader
                                    While rst.Read
                                        ReportColHeader += "<td>" & rst("Field").ToString & "</td>"
                                        tmpField += rst("Field").ToString & ","
                                        cnt += 1
                                    End While
                                    rst.Close()

                                    If tmpField.Length > 0 Then
                                        tmpField = tmpField.Remove(tmpField.Length - 1, 1)
                                        ReportHeader = "<table border='0' id='tabl_sub_border' style='margin-bottom:20px;' cellspacing='0' cellpadding='0'>" & Environment.NewLine _
                                                                + "<tr align='center' style='font-weight: bold;'><td align='center' colspan='" & cnt & "' style='padding: 10px 0; background-color: #" & If(.Rows(x)("bgcolor").ToString() = "", "FFF", .Rows(x)("bgcolor").ToString()) & "; color: #" & If(.Rows(x)("forecolor").ToString() = "", "000", .Rows(x)("forecolor").ToString()) & ";'>" & ReplaceDateStr(.Rows(x)("rpttitle").ToString(), ReportDate) & "</td></tr>" & Environment.NewLine _
                                                                + "<tr align='center' style='font-weight: bold;'>" & ReportColHeader & "</tr>" & Environment.NewLine

                                        com.CommandText = "select * from " & .Rows(x)("shortcode").ToString() & "" : rst = com.ExecuteReader
                                        While rst.Read
                                            Dim ColumnData As String = "" : Dim foundSummaryvalue As Boolean = False
                                            For Each ColumnName In tmpField.Split(New Char() {","c})
                                                Dim ColAlign As String = "" : Dim ColValue As String = ""
                                                If CenterColumn(ColumnName) = True Then
                                                    ColAlign = "align='center'"
                                                    ColValue = rst(ColumnName).ToString
                                                Else
                                                    If RightColumn(ColumnName) = True Then
                                                        ColAlign = "align='right'"
                                                        ColValue = FormatNumber(Val(rst(ColumnName).ToString), 2)
                                                    Else
                                                        ColValue = rst(ColumnName).ToString
                                                    End If
                                                End If
                                                ColumnData += "<td " & ColAlign & ">" & ColValue & "</td>"

                                                If rst(ColumnName).ToString.Contains("TOTAL SUMMARY") = True Then
                                                    foundSummaryvalue = True
                                                End If
                                            Next
                                            ReportRowValue += "<tr " & If(foundSummaryvalue = True, "style='font-weight: bold;'", "") & ">" & ColumnData & "</tr>" & Environment.NewLine
                                        End While
                                        rst.Close()
                                        ReportTable = ReportHeader + ReportRowValue & "</table>" & Environment.NewLine
                                    End If
                                End If
                                
                                HtmlBody += If(ReportTable.Length > 2, ReportTable & Environment.NewLine, "")
                            End With
                        Next

                        If GlobalOrganizationLogoURL.Length > 5 Then
                            My.Computer.FileSystem.WriteAllText(SaveLocation, My.Computer.FileSystem.ReadAllText(SaveLocation).Replace("[logo]", "<div align='center'><img src='" & GlobalOrganizationLogoURL & "'></div>"), False)
                        Else
                            My.Computer.FileSystem.WriteAllText(SaveLocation, My.Computer.FileSystem.ReadAllText(SaveLocation).Replace("[logo]", ""), False)
                        End If

                        My.Computer.FileSystem.WriteAllText(SaveLocation, My.Computer.FileSystem.ReadAllText(SaveLocation).Replace("[title]", ReportTitle), False)
                        My.Computer.FileSystem.WriteAllText(SaveLocation, My.Computer.FileSystem.ReadAllText(SaveLocation).Replace("[details]", ReportDetails), False)
                        My.Computer.FileSystem.WriteAllText(SaveLocation, My.Computer.FileSystem.ReadAllText(SaveLocation).Replace("[transaction]", HtmlBody), False)
                        Dim EmailReport As String = My.Computer.FileSystem.ReadAllText(SaveLocation)

                        If EmailReport.Length > 0 Then
                            If reportType = "Weekly" Then
                                InsertHTMLEmailNotification("SUMMARY", EmailList, StrConv(ReportTitle & " " & ReportDate, vbProperCase), EmailReport, "UPDATE tblgeneralsettings set weeklyreportdate='" & ConvertDate(CDate(ReportDate).AddDays(7)) & "';")
                            ElseIf reportType = "Monthly" Then
                                InsertHTMLEmailNotification("SUMMARY", EmailList, StrConv(ReportTitle & " " & CDate(ReportDate).AddMonths(-1).ToString("MMMM yyyy"), vbProperCase), EmailReport, "UPDATE tblgeneralsettings set monthlyreportdate='" & ConvertDate(CDate(ReportDate).AddMonths(1)) & "';")
                            End If
                        End If
                    End If
                End If

            End If
        Catch ex As Exception
            RecordLog("Error Notification: " + ex.Message)
        End Try
        Return True
    End Function

    Public Sub InsertHTMLEmailNotification(ByVal trntype As String, ByVal receiver As String, ByVal subject As String, emailbody As String, ByVal CommandQuery As String)
        If receiver.Length > 5 Then
            com.CommandText = "insert into tblemailnotification set trntype='" & trntype & "', replyto='', receiver='" & receiver & "', subject='" & Trim(rchar(subject)) & "', emailbody='" & EncryptTripleDES(FormattingEmailBody(emailbody)) & "'" : com.ExecuteNonQuery()
        End If
        If CommandQuery <> "" Then
            com.CommandText = CommandQuery : com.ExecuteNonQuery()
        End If
    End Sub
     
    Public Function FormattingEmailBody(ByVal value As String) As String
        value = value.Replace("Ñ", "&Ntilde;")
        value = value.Replace("ñ", "&ntilde;")
        value = value.Replace("  -  ", "&nbsp; &nbsp; - &nbsp; &nbsp;")
        Return value
    End Function
    Public Function rchar(ByVal str As String)
        str = str.Replace("'", "''")
        str = str.Replace("\", "\\")
        Return str
    End Function

    Public Function CenterColumn(ByVal str As String) As Boolean
        CenterColumn = False
        If LCase(str).Contains("unit") Then
            CenterColumn = True
        ElseIf LCase(str).Contains("quantity") Then
            CenterColumn = True
        ElseIf LCase(str).Contains("date") Then
            CenterColumn = True
        ElseIf LCase(str).Contains("product id") Then
            CenterColumn = True
        ElseIf LCase(str).Contains("birth date") Then
            CenterColumn = True
        ElseIf LCase(str).Contains("age") Then
            CenterColumn = True
        ElseIf LCase(str).Contains("office") Then
            CenterColumn = True
        ElseIf LCase(str).Contains("branch") Then
            CenterColumn = True
        ElseIf LCase(str).Contains("time") Then
            CenterColumn = True
        ElseIf LCase(str).Contains("transaction") Then
            CenterColumn = True
        ElseIf LCase(str).Contains("contact") Then
            CenterColumn = True
        End If

        Return CenterColumn
    End Function

    Public Function RightColumn(ByVal str As String) As Boolean
        RightColumn = False
        If LCase(str).Contains("total") Then
            RightColumn = True
        ElseIf LCase(str).Contains("gross") Then
            RightColumn = True
        ElseIf LCase(str).Contains("revenue") Then
            RightColumn = True
        ElseIf LCase(str).Contains("income") Then
            RightColumn = True
        ElseIf LCase(str).Contains("total amount") Then
            RightColumn = True
        ElseIf LCase(str).Contains("amount") Then
            RightColumn = True
        ElseIf LCase(str).Contains("summary") Then
            RightColumn = True
        ElseIf LCase(str).Contains("cost") Then
            RightColumn = True
        ElseIf LCase(str).Contains("payment") Then
            RightColumn = True
        ElseIf LCase(str).Contains("deposit") Then
            RightColumn = True
        ElseIf LCase(str).Contains("balance") Then
            RightColumn = True
        End If

        Return RightColumn
    End Function

    Public Function ReplaceDateStr(ByVal source As String, ByVal paramDate As String) As String
        Dim str = source
        str = str.Replace("[date]", CDate(paramDate).AddDays(-1).ToString("MMMM dd, yyyy"))
        str = str.Replace("[day]", CDate(paramDate).AddDays(-1).ToString("dd"))
        str = str.Replace("[month]", CDate(paramDate).AddDays(-1).ToString("MMMM"))
        str = str.Replace("[year]", CDate(paramDate).AddDays(-1).ToString("yyyy"))
        Return str
    End Function
End Module
