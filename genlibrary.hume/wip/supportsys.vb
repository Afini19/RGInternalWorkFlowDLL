Imports Microsoft.VisualBasic
Imports System.Data.OleDB
Imports System.Data
Public Class SupportPreference

    Public Shared Property SMTPServerCode()
        Get
            If (HttpContext.Current.Session("SMTPServerCode") & "" <> "") Then
                Return HttpContext.Current.Session("SMTPServerCode")
            Else
                Return ""
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("SMTPServerCode") = value

        End Set
    End Property
    Public Shared Property FromNAme()
        Get
            If (HttpContext.Current.Session("FromNAme") & "" <> "") Then
                Return HttpContext.Current.Session("FromNAme")
            Else
                Return ""
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("FromNAme") = value

        End Set
    End Property
    Public Shared Property FromEmail()
        Get
            If (HttpContext.Current.Session("FromEmail") & "" <> "") Then
                Return HttpContext.Current.Session("FromEmail")
            Else
                Return ""
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("FromEmail") = value

        End Set
    End Property
    Public Shared Function LoadSettings() As Boolean
        If SettingsLoaded = False Then
            If getsettings() = True Then
                SettingsLoaded = True
                Return True
            Else
                Return False
            End If
        Else
            Return True
        End If
    End Function
    Public Shared Property SettingsLoaded() As Boolean
        Get
            If (HttpContext.Current.Session("SettingsLoaded") & "" <> "") Then
                Return HttpContext.Current.Session("SettingsLoaded")
            Else
                Return False
            End If
        End Get
        Set(ByVal value As Boolean)
            HttpContext.Current.Session("SettingsLoaded") = value

        End Set
    End Property
    Public Shared Property SendSubmitEmail() As Boolean
        Get
            If (HttpContext.Current.Session("SendSubmitEmail") & "" <> "") Then
                Return HttpContext.Current.Session("SendSubmitEmail")
            Else
                Return ""
            End If
        End Get
        Set(ByVal value As Boolean)
            HttpContext.Current.Session("SendSubmitEmail") = value

        End Set
    End Property
    Public Shared Property SendCloseEmail() As Boolean
        Get
            If (HttpContext.Current.Session("SendCloseEmail") & "" <> "") Then
                Return HttpContext.Current.Session("SendCloseEmail")
            Else
                Return ""
            End If
        End Get
        Set(ByVal value As Boolean)
            HttpContext.Current.Session("SendCloseEmail") = value

        End Set
    End Property
    Public Shared Property SendReplyEmail() As Boolean
        Get
            If (HttpContext.Current.Session("SendReplyEmail") & "" <> "") Then
                Return HttpContext.Current.Session("SendReplyEmail")
            Else
                Return ""
            End If
        End Get
        Set(ByVal value As Boolean)
            HttpContext.Current.Session("SendReplyEmail") = value

        End Set
    End Property
    Public Shared Function getSettings() As Boolean

        Dim connectionstring As String = System.Configuration.ConfigurationSettings.AppSettings("ConnStr")
        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim dr As DataRow
        Dim ltemp As String = ""
        Dim lDocNo As String = ""
        Try


            cn.open()
            'cmd.CommandText = "Select top 1 * from SupportPref where sp_merchantid='" & Weblib.Merchantid & "' and sp_filter='" & Weblib.Filtercode & "' order by sp_id asc"
            cmd.CommandText = "Select top 1 * from SupportPref order by sp_id asc"

            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")
            For Each dr In ds.Tables("datarecords").Rows
                counter = counter + 1
                SMTPServerCode = dr("sp_smtpserver") & ""
                SendReplyEmail = Weblib.bittoboolean(dr("sp_sendemailreply") & "")
                SendSubmitEmail = Weblib.bittoboolean(dr("sp_sendemailonsubmit") & "")
                SendCloseEmail = Weblib.bittoboolean(dr("sp_sendmailonclose") & "")
                FromNAme = dr("sp_emailfromname") & ""
                FromEmail = dr("sp_emailfromemail") & ""
                Exit For

            Next
            cn.Close()
            cmd.dispose()
            cn.dispose()

            If FromName.trim = "" Then
                Weblib.ErrorTrap = "From Name Not Set"
                Return False
                Exit Function
            End If
            If FromEmail.trim = "" Then
                Weblib.ErrorTrap = "From Email Not Set"
                Return False
                Exit Function
            End If

            If SMTPServerCode.trim = "" Then
                Weblib.ErrorTrap = "SMTP Server Not Set"
                Return False
                Exit Function
            End If


            If counter = 0 Then
                Weblib.ErrorTrap = "Support Preference Not Set"
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            Weblib.ErrorTrap = ex.Message

            Return False
        End Try
    End Function
    Public Shared Function ReplaceEmailData(ByVal TicketNo As String, ByVal pReplyMessage As String, ByRef pMessage As String, ByRef pSubject As String) As String

        Dim connectionstring As String = System.Configuration.ConfigurationSettings.AppSettings("ConnStr")
        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim dr As DataRow
        Try


            cn.open()
            'cmd.CommandText = "Select top 1 * from SupportTicket where tic_no='" & TicketNo & "'  and tic_merchantid='" & Weblib.Merchantid & "' and tic_filter='" & Weblib.Filtercode & "'"
            cmd.CommandText = "Select top 1 * from SupportTicket where tic_no='" & TicketNo & "'"

            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")
            For Each dr In ds.Tables("datarecords").Rows
                counter = counter + 1
                pMessage = pMessage.replace("#TicketNo#", dr("tic_no") & "")
                pMessage = pMessage.replace("#Subject#", dr("tic_subject") & "")
                pMessage = pMessage.replace("#TicketDetails#", dr("tic_details") & "")
                pMessage = pMessage.replace("#SubmitDate#", dr("tic_createdt") & "")
                pMessage = pMessage.replace("#Status#", dr("tic_status") & "")
                pMessage = pMessage.replace("#ReplyMessage#", pReplyMessage)

                pSubject = pSubject.replace("#TicketNo#", dr("tic_no") & "")
                pSubject = pSubject.replace("#Subject#", dr("tic_subject") & "")
                pSubject = pSubject.replace("#TicketDetails#", dr("tic_details") & "")
                pSubject = pSubject.replace("#SubmitDate#", dr("tic_createdt") & "")
                pSubject = pSubject.replace("#Status#", dr("tic_status") & "")
                pSubject = pSubject.replace("#ReplyMessage#", pReplyMessage)

                Dim lAgentNAme As String = Weblib.GetValue("secuserinfo", "usr_name", "usr_loginid", "'" & dr("tic_assignto") & "'", "usr_merchantid", "usr_filter")
                Dim lUserNAme As String = Weblib.GetValue("secuserinfo", "usr_name", "usr_loginid", "'" & dr("tic_createby") & "'", "usr_merchantid", "usr_filter")
                pMessage = pMessage.replace("#SupportAgentName#", lAgentNAme)
                pMessage = pMessage.replace("#SubmitByName#", lUserNAme)

                pSubject = pSubject.replace("#SupportAgentName#", lAgentNAme)
                pSubject = pSubject.replace("#SubmitByName#", lUserNAme)


                Exit For
            Next

            cn.Close()
            cmd.dispose()
            cn.dispose()

            If counter = 0 Then
                Return pMessage
            Else
                Return pMessage
                Return True
            End If
        Catch ex As Exception
            Return pMessage
        End Try
    End Function
    Public Shared Function EmailFields() As String
        Dim lFieldName As String = ""
        lFieldName = lFieldName & "#TicketNo#<br>"
        lFieldName = lFieldName & "#Subject#<br>"
        lFieldName = lFieldName & "#TicketDetails#<br>"
        lFieldName = lFieldName & "#SubmitDate#<br>"
        lFieldName = lFieldName & "#SupportAgentName#<br>"
        lFieldName = lFieldName & "#ReplyMessage#<br>"
        lFieldName = lFieldName & "#SubmitByName#<br>"
        lFieldName = lFieldName & "#Status#<br>"
    End Function

End Class

Public Class clsSupportEmail
    Public ErrorMsg As String = ""
    Public connectionstring As String = System.Configuration.ConfigurationSettings.AppSettings("ConnStr")
    Public ReadOnly Property GetWorflowCustomFields()
        Get
            Return "#TicketNo#|#Subject#|#TicketDetails#|#SubmitDate#|#SupportAgentName#|#ReplyMessage#|#SubmitByName#|#Status#"
        End Get
    End Property
    Public Function NotifyUsers(ByVal TemplateCode As String, ByVal ReferenceData As String) As Boolean
        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim dr As DataRow
        Dim orgEmail As String
        Dim orgSubject As String
        Dim objEmail As New EmailObject
        Dim lTemplate As String = ""
        Dim lFromName As String = ""  'Pending Assignment
        Dim lDepartment As String = "" 'Pending Assignment

        If (templatecode & "").trim = "" Then
            Return False
            Exit Function
        End If

        If SupportPreference.LoadSettings = True And SupportPreference.SendSubmitEmail = True Then
            Dim objEmail As New EmailObject

            If objEmail.GetEmailTemplate(TemplateCode, SupportPreference.SMTPServerCode) = True Then

                SupportPreference.ReplaceEmailData(tic_no.text, "", objEmail.EmailBody, objEmail.EmailSubject)
                objEmail.Emailto = WebLib.LoginUserEmail
                objEmail.FromName = SupportPreference.FromName
                objEmail.FromEmail = SupportPreference.FromEmail
                If objEmail.sendEmail() = False Then
                    ErrorMsg = objEmail.ErrorMsg
                    Return False
                End If
            Else
                ErrorMsg = "Enable to Retrieve Email Format"
                Return False
            End If
        Else
            ErrorMsg = "Support Preference Settings Not Set"
            Return False
            Exit Function
        End If



        If objEmail.GetEmailTemplate(lTemplate, "GENERAL") = True Then
            orgSubject = objEmail.EmailSubject
            orgEmail = objEmail.EmailBody
        Else
            ErrorMsg = "Error Init Sending"
            Return False
            Exit Function
        End If

        

        Try
            Call EmailFieldsReplace(ReferenceData, orgSubject, orgEmail)
        Catch ex As Exception

        End Try

        Try
            cmd.CommandText = "Select usr_hpno,usr_hpcc from supportstaffdept inner join secuserinfo on sd_userid = usr_loginid and sd_merchantid=usr_merchantid and sd_filter = usr_filter where sd_deptid=" & lDepartmentID
            If (cmd.commandtext & "").trim = "" Then
                ErrorMsg = "Unable to retrieve query"
                Return False
                Exit Function
            End If
            cn.open()
            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")


            For Each dr In ds.Tables("datarecords").Rows

                counter = counter + 1
                objEmail.EmailSubject = orgSubject
                objEmail.EmailBody = orgEmail

                objEmail.Emailto = dr("usr_email") & ""
                objEmail.FromName = lFromName
                If objEmail.sendEmail() = False Then
                    ErrorMsg = Weblib.ErrorTrap
                    Return False
                    Exit Function
                End If

            Next

            cn.Close()
            cmd.Dispose()
            cn.Dispose()
            Return True
        Catch ex As Exception
            ErrorMSg = ex.MEssage
            Return False
        End Try
    End Function
    Private Function EmailFieldsReplace(ByVal ReferenceData As String, ByRef EmailSubject As String, ByRef Emailbody As String) As Boolean
        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim dr As DataRow

        Try
            cn.open()
            cmd.CommandText = "Select top 1 * from SupportTicket where tic_no='" & ReferenceData & "'"

            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")
            For Each dr In ds.Tables("datarecords").Rows

                Emailbody = Emailbody.replace("#TicketNo#", dr("tic_no") & "")
                Emailbody = Emailbody.replace("#Subject#", dr("tic_subject") & "")
                Emailbody = Emailbody.replace("#TicketDetails#", dr("tic_details") & "")
                Emailbody = Emailbody.replace("#SubmitDate#", dr("tic_createdt") & "")
                Emailbody = Emailbody.replace("#Status#", dr("tic_status") & "")
                Emailbody = Emailbody.replace("#ReplyMessage#", pReplyMessage)

                EmailSubject = EmailSubject.replace("#TicketNo#", dr("tic_no") & "")
                EmailSubject = EmailSubject.replace("#Subject#", dr("tic_subject") & "")
                EmailSubject = EmailSubject.replace("#TicketDetails#", dr("tic_details") & "")
                EmailSubject = EmailSubject.replace("#SubmitDate#", dr("tic_createdt") & "")
                EmailSubject = EmailSubject.replace("#Status#", dr("tic_status") & "")
                EmailSubject = EmailSubject.replace("#ReplyMessage#", pReplyMessage)

                Dim lAgentNAme As String = Weblib.GetValue("secuserinfo", "usr_name", "usr_loginid", "'" & dr("tic_assignto") & "'", "usr_merchantid", "usr_filter")
                Dim lUserNAme As String = Weblib.GetValue("secuserinfo", "usr_name", "usr_loginid", "'" & dr("tic_createby") & "'", "usr_merchantid", "usr_filter")
                pMessage = pMessage.replace("#SupportAgentName#", lAgentNAme)
                pMessage = pMessage.replace("#SubmitByName#", lUserNAme)
                pSubject = pSubject.replace("#SupportAgentName#", lAgentNAme)
                pSubject = pSubject.replace("#SubmitByName#", lUserNAme)

                Exit For
            Next
            cn.Close()
            cmd.Dispose()
            cn.Dispose()

            Return True

        Catch ex As Exception
            Return False
        End Try
    End Function



End Class