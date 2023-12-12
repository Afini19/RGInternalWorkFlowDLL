Imports System
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Configuration
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.HtmlControls
Imports System.Text
Imports System.Collections.Generic
Imports System.Xml
Imports System.Text.RegularExpressions

Public Class SOSettings
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

            ltemp = ""

            cn.open()
            cmd.CommandText = "Select top 1 * from sysSettings"
            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")
            For Each dr In ds.Tables("datarecords").Rows
                counter = counter + 1
                NeedDateMinDays = dr("ss_ndmin") & ""
                NeedDateMaxDays = dr("ss_ndmax") & ""
                OrderCutOffTime = dr("ss_cutofftime") & ""
                CustomerProfile = dr("ss_profile") & ""
                CustomerBranch = dr("ss_branch") & ""
                DocPath1 = dr("ss_docpath1") & ""
                DocPath2 = dr("ss_docpath2") & ""

            Next
            cn.Close()
            cmd.dispose()
            cn.dispose()

            If counter = 0 Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            Return False
        End Try


    End Function


    Public Shared Function getShipToSettings(ByRef Subject As String, ByRef EmailFromNAme As String, ByRef EmailCC As String, ByRef EmailBCC As String, ByRef EmailMessage As String, ByRef SMTPCode As String) As Boolean
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

            ltemp = ""

            cn.open()
            cmd.CommandText = "Select top 1 * from sysSettings"
            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")
            For Each dr In ds.Tables("datarecords").Rows
                counter = counter + 1
                EmailFromNAme = dr("ss_stfromname") & ""
                EmailCC = dr("ss_stccemail") & ""
                EmailBCC = dr("ss_stbccemail") & ""
                Subject = dr("ss_stsubject") & ""
                EmailMessage = ""
                SMTPCode = "GENERAL"

            Next
            cn.Close()
            cmd.dispose()
            cn.dispose()

            If counter = 0 Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            Return False
        End Try


    End Function


    Public Shared Property NeedDateMinDays()
        Get
            If (HttpContext.Current.Session("NeedDateMinDays") & "" <> "") Then
                If isnumeric(HttpContext.Current.Session("NeedDateMinDays")) = False Then
                    Return 0
                Else
                    Return HttpContext.Current.Session("NeedDateMinDays")
                End If
            Else
                Return HttpContext.Current.Session("NeedDateMinDays")
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("NeedDateMinDays") = value

        End Set

    End Property
    Public Shared Function InitRuntimeObject() As Boolean
        Try

            If (NeedDateMinDays & "").trim = "" Then
                If getSettings() = False Then
                    Return False
                    Exit Function
                End If
            End If
            Return True

        Catch ex As Exception
            Return False
        End Try

    End Function

    Public Shared Property NeedDateMaxDays()
        Get
            If (HttpContext.Current.Session("NeedDateMaxDays") & "" <> "") Then

                If isnumeric(HttpContext.Current.Session("NeedDateMaxDays")) = False Then
                    Return 0
                Else
                    Return HttpContext.Current.Session("NeedDateMaxDays")
                End If
            Else
                Return "0"
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("NeedDateMaxDays") = value

        End Set

    End Property
    Public Shared Property DocPath2()
        Get
            If (HttpContext.Current.Session("DocPath2") & "" <> "") Then
                Return HttpContext.Current.Session("DocPath2")
            Else
                Return ""
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("DocPath2") = value

        End Set

    End Property

    Public Shared Property DocPath1()
        Get
            If (HttpContext.Current.Session("DocPath1") & "" <> "") Then
                Return HttpContext.Current.Session("DocPath1")
            Else
                Return ""
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("DocPath1") = value

        End Set

    End Property
    Public Shared Property CustomerBranch()
        Get
            If (HttpContext.Current.Session("CustomerBranch") & "" <> "") Then
                Return HttpContext.Current.Session("CustomerBranch")
            Else
                Return HttpContext.Current.Session("CustomerBranch")
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("CustomerBranch") = value

        End Set

    End Property

    Public Shared Property CustomerProfile()
        Get
            If (HttpContext.Current.Session("CustomerProfile") & "" <> "") Then
                Return HttpContext.Current.Session("CustomerProfile")
            Else
                Return HttpContext.Current.Session("CustomerProfile")
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("CustomerProfile") = value

        End Set

    End Property
    Public Shared Property OrderCutOffTime()
        Get
            If (HttpContext.Current.Session("OrderCutOffTime") & "" <> "") Then
                Return HttpContext.Current.Session("OrderCutOffTime")
            Else
                Return HttpContext.Current.Session("OrderCutOffTime")
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("OrderCutOffTime") = value

        End Set

    End Property

End Class

Public Class POSSettings
    Public Shared Function getPOSSettingsbyBranch() As Boolean

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

            ltemp = ltemp & " where poss_merchantid='" & WebLib.MerchantID & "'"
            ltemp = ltemp & " and poss_filter='" & WebLib.FilterCode & "'"
            ltemp = ltemp & " and poss_branchid=" & WebLib.Branchid & ""

            cn.open()
            cmd.CommandText = "Select * from possettings " & ltemp
            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")
            For Each dr In ds.Tables("datarecords").Rows
                counter = counter + 1
                ForceResource1 = dr("poss_forceresource1")
                ForceResource2 = dr("poss_forceresource2")
                EnableFastCheck = dr("poss_fastcheck")
                QueueResourceCat = dr("poss_resourcequeue")
                Branchid = dr("poss_branchid")
                EnableResource1 = dr("poss_resource1")
                Resource1Cat = dr("poss_resource1cat")
                Resource1Desc = dr("poss_ressource1desc")
                EnableResource2 = dr("poss_resource2")
                Resource2Cat = dr("poss_resource2cat")
                Resource2Desc = dr("poss_ressource2desc")


                If isdate(dr("poss_businessdate") & "") = True Then
                    BusinessDate = Weblib.Formatthedate(dr("poss_businessdate"))
                    BusinessDateDate = dr("poss_businessdate")
                Else
                    BusinessDate = Weblib.Formatthedate(datetime.today)
                    BusinessDateDate = datetime.today

                    cmd.commandtext = "Update possettings set poss_businessdate='" & BusinessDate & "' " & ltemp
                    cmd.ExecuteNonQuery()

                End If


            Next
            cn.Close()
            cmd.dispose()
            cn.dispose()

            If counter = 0 Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Shared Property ForceResource1()
        Get
            If (HttpContext.Current.Session("ForceResource1") & "" <> "") Then
                Return HttpContext.Current.Session("ForceResource1")
            Else
                Return ""
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("ForceResource1") = value

        End Set

    End Property
    Public Shared Property BusinessDate()
        Get
            If (HttpContext.Current.Session("BusinessDate") & "" <> "") Then
                Return HttpContext.Current.Session("BusinessDate")
            Else
                Return ""
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("BusinessDate") = value

        End Set
    End Property
    Public Shared Property BusinessDateDate()
        Get
            If (HttpContext.Current.Session("BusinessDateDate") & "" <> "") Then
                Return HttpContext.Current.Session("BusinessDateDate")
            Else
                Return ""
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("BusinessDateDate") = value

        End Set
    End Property
    Public Shared Property ForceResource2()
        Get
            If (HttpContext.Current.Session("ForceResource2") & "" <> "") Then
                Return HttpContext.Current.Session("ForceResource2")
            Else
                Return ""
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("ForceResource2") = value

        End Set
    End Property
    Public Shared Property EnableFastCheck()
        Get
            If (HttpContext.Current.Session("EnableFastCheck") & "" <> "") Then
                Return HttpContext.Current.Session("EnableFastCheck")
            Else
                Return ""
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("EnableFastCheck") = value

        End Set
    End Property
    Public Shared Property QueueResourceCat()
        Get
            If (HttpContext.Current.Session("QueueResourceCat") & "" <> "") Then
                Return HttpContext.Current.Session("QueueResourceCat")
            Else
                Return ""
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("QueueResourceCat") = value

        End Set
    End Property

    Public Shared Property Branchid()
        Get
            If (HttpContext.Current.Session("possetbranchid") & "" <> "") Then
                Return HttpContext.Current.Session("possetbranchid")
            Else
                Return ""
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("possetbranchid") = value

        End Set
    End Property
    Public Shared Property EnableResource1()
        Get
            If (HttpContext.Current.Session("EnableResource1") & "" <> "") Then
                Return HttpContext.Current.Session("EnableResource1")
            Else
                Return ""
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("EnableResource1") = value

        End Set
    End Property
    Public Shared Property Resource1Cat()
        Get
            If (HttpContext.Current.Session("Resource1Cat") & "" <> "") Then
                Return HttpContext.Current.Session("Resource1Cat")
            Else
                Return ""
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("Resource1Cat") = value

        End Set
    End Property
    Public Shared Property Resource1Desc()
        Get
            If (HttpContext.Current.Session("Resource1Desc") & "" <> "") Then
                Return HttpContext.Current.Session("Resource1Desc")
            Else
                Return ""
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("Resource1Desc") = value

        End Set
    End Property
    Public Shared Property EnableResource2()
        Get
            If (HttpContext.Current.Session("EnableResource2") & "" <> "") Then
                Return HttpContext.Current.Session("EnableResource2")
            Else
                Return ""
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("EnableResource2") = value

        End Set
    End Property
    Public Shared Property Resource2Cat()
        Get
            If (HttpContext.Current.Session("Resource2Cat") & "" <> "") Then
                Return HttpContext.Current.Session("Resource2Cat")
            Else
                Return ""
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("Resource2Cat") = value

        End Set
    End Property
    Public Shared Property Resource2Desc()
        Get
            If (HttpContext.Current.Session("Resource2Desc") & "" <> "") Then
                Return HttpContext.Current.Session("Resource2Desc")
            Else
                Return ""
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("Resource2Desc") = value

        End Set
    End Property
End Class

Public Class EmailObject
    Public SmtpServer As String
    Public EmailSubject As String
    Public EmailBody As String
    Public Emailto As String
    Public Emailcc As String = System.Configuration.ConfigurationSettings.AppSettings("WfEmailCC")
    Public Emailbcc As String = System.Configuration.ConfigurationSettings.AppSettings("WfEmailBCC")

    Public FromName As String
    Public FromEmail As String

    Public _smtpserver As String
    Public _smtpport As String
    Public _LoginID As String
    Public _FromEmail As String
    Public _SSL As Boolean = False
    Public _LoginPassword As String
    Public _SMTPDEfined As Boolean = False
    Private connectionstring As String = System.Configuration.ConfigurationSettings.AppSettings("ConnStr")
    Public Function loadSMTP(ByVal pSMTPCode As String) As Boolean


        Dim cn As New OledbConnection(connectionstring)
        Dim cmd As New OledbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim dr As DataRow

        Try
            cn.open()
            'cmd.commandText = "Select * from sysSMTP where smtp_type='" & pSMTPCode & "' and smtp_merchantid='" & Weblib.MerchantID & "' and smtp_filter='" & Weblib.filtercode & "'"
            cmd.commandText = "Select * from sysSMTP where smtp_type='" & pSMTPCode & "'"

            cmd.connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "Registration")
            For Each dr In ds.Tables("Registration").Rows
                counter = counter + 1
                _smtpserver = dr("smtp_server") & ""
                _smtpport = dr("smtp_port") & ""
                _LoginID = dr("smtp_login") & ""
                _LoginPassword = dr("smtp_password") & ""
                _FromEmail = dr("smtp_fromemail") & ""
                _SSL = weblib.bittoboolean(dr("smtp_ssl") & "")
                FromEmail = dr("smtp_fromemail") & ""
                _SMTPDEfined = True
                Exit For
            Next

            If counter = 0 Then
                Weblib.ErrorTrap = "Unable to Load SMTP Settings"
                _SMTPDEfined = False
                Return False
            End If

            cn.close()
            Return True

        Catch ex As Exception
            Weblib.ErrorTrap = ex.message
            Return False
        End Try
    End Function
    Public Function sendEmail() As Boolean

        Dim lbody As String
        Try
            lbody = "<html><body><font style=""color:black; font-family:Arial; font-size:12px"">" & EmailBody & "</font></body></html>"
            '            Dim sendmail As New VIEmailAccess.Email2u
            Dim sendmail As New VIFeandiEmail.Email2u



            sendmail.InitSMTPServer(_smtpserver, FromEmail, _smtpport, _LoginID, _LoginPassword)


            sendmail.SendEmail(Emailto, FromEmail, Emailto, FromName, EmailSubject, lbody, True, _SSL, Emailcc, Emailbcc)


            sendmail = Nothing
            Return True

        Catch ex As Exception
            Weblib.ErrorTrap = ex.message & _smtpserver & FromEmail & _smtpport & _LoginID & _LoginPassword & Emailto
            Return False
        End Try

    End Function
    Public Function GetEmailTemplate(ByVal pEmailTypeCode As String, ByVal pEmailSMTPCode As String) As Boolean

        Dim cn As New OledbConnection(connectionstring)
        Dim cmd As New OledbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim dr As DataRow
        Dim lPayType As String = ""
        Dim lSubject As String = ""
        Dim lBody As String = ""
        Dim lFrom As String
        Dim lfooter As String = ""
        Dim lTableName As String = "EmailTemplate"


        '        Try
        'cmd.commandText = "Select * from " & lTableName & " where email_type='" & pEmailTypeCode & "' and email_merchantid='" & Weblib.MerchantID & "' and email_filter='" & Weblib.filtercode & "'"
        cmd.commandText = "Select * from " & lTableName & " where email_type='" & pEmailTypeCode & "'"

retryagain:

        cmd.connection = cn
        ad.SelectCommand = cmd
        ad.Fill(ds, "Registration")
        For Each dr In ds.Tables("Registration").Rows
            counter = counter + 1
            Emailsubject = dr("email_subject") & ""
            EmailBody = dr("email_body") & ""
            Exit For
        Next

        If counter = 0 Then

            If lTableName <> "EmailTemplateDefault" Then
                lTableName = "EmailTemplateDefault"
                cmd.commandText = "Select * from " & lTableName & " where email_type='" & pEmailTypeCode & "'"

                GoTo retryagain
            End If

            Weblib.ErrorTrap = "There is no Email Template Set "
            Return False
        End If
        cn.close()

        If (pEmailSMTPCode & "").trim = "" Then
            Weblib.ErrorTrap = "SMTP Code Not Set"
            Return False
            Exit Function
        End If
        If loadSMTP(pEmailSMTPCode) = False Then
            Return False
            Exit Function
        End If

        Return True
        '        Catch ex As Exception
        'Weblib.ErrorTrap = ex.message
        ' Return False
        'End Try

    End Function

End Class

Public Class EmailSettings
    Private lSMTPCode As String
    Private lFromName As String
    Private lFromEmail As String
    Public Property SMTPServerCode()
        Get
            Return lSMTPCode
        End Get
        Set(ByVal value)
            lSMTPCode = value
        End Set
    End Property
    Public Property FromNAme()
        Get
            Return lFromName
        End Get
        Set(ByVal value)
            lFromName = value
        End Set
    End Property
    Public Property FromEmail()
        Get
            Return lFromEmail
        End Get
        Set(ByVal value)
            lFromEmail = value
        End Set
    End Property
    Public Function getSettings(ByVal pModuleCode As String, ByVal pMerchantID As String, ByVal pFilterCode As String) As Boolean

        Dim connectionstring As String = System.Configuration.ConfigurationSettings.AppSettings("ConnStr")
        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim dr As DataRow
        Dim ltemp As String = ""
        Try


            cn.open()
            'cmd.CommandText = "Select top 1 * from sysSMTPSet where sss_modcode='" & pModuleCode & "' and isnull(sss_merchantid,'')='" & pMerchantID & "' and isnull(sss_filter,'')='" & pFilterCode & "'"
            cmd.CommandText = "Select top 1 * from sysSMTPSet where sss_modcode='" & pModuleCode & "'"
            
            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")
            For Each dr In ds.Tables("datarecords").Rows
                counter = counter + 1
                SMTPServerCode = dr("sss_smtpcode") & ""
                FromNAme = dr("sss_fromname") & ""
                FromEmail = dr("sss_fromemail") & ""
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
                Weblib.ErrorTrap = "SMTP Settings Not Set"
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            Weblib.ErrorTrap = ex.Message

            Return False
        End Try
    End Function
End Class


Public Class NotificationServerGeneral
    Public ErrorMsg As String = ""
    Private connectionstring As String = System.Configuration.ConfigurationSettings.AppSettings("ConnStr")
    Public Function NotifyEmail(ByVal EventCode As String, ByVal SMTPCode As String, ByVal EmailSubject As String, ByVal EmailMessage As String, ByVal EmailFromName As String, ByVal EmailTo As String, ByVal EmailCC As String, ByVal EmailBcc As String, ByVal DirectSend As Boolean, Optional ByVal EmailToName As String = "") As Boolean
        Dim counter As Integer = 0
        Dim objEmail As New EmailObject

        Try

            If objEmail.loadSMTP(SMTPCode) = False Then
                ErrorMsg = "Error Init Sending"
                Return False
                Exit Function
            End If


            objEmail.EmailSubject = EmailSubject
            objEmail.EmailBody = EmailMessage

            objEmail.Emailto = EmailTo
            objEmail.FromName = EmailFromName

            objEmail.Emailcc = emailcc
            objEmail.Emailbcc = emailbcc
            If DirectSend = True Then
                If objEmail.sendEmail() = False Then
                    ErrorMsg = Weblib.ErrorTrap
                    Return False
                    Exit Function
                End If
            End If


            '************ NOTIFICATION CENTER ********************
            Dim ooonoti As New ViFeandi.NotificationCenter

            Dim ConnectionStr As String = connectionstring
            Dim noti_MerchantID As String = Weblib.MerchantID
            Dim noti_Message As String = objEmail.EmailBody
            Dim noti_CCCode As String = ""
            Dim noti_MobileNo As String = ""
            Dim noti_EventCode As String = EVENTCODE


            Dim noti_toNAme As String
            If (EmailToName & "") <> "" Then
                noti_toname = EmailToName
            Else
                noti_toname = objEmail.EmailTo
            End If


            Dim noti_toEmail As String = objEmail.EmailTo
            Dim noti_FromNAme As String = objEmail.FromName
            Dim noti_FromEmail As String = objEmail._FromEmail
            Dim noti_EmailSubject As String = objEmail.EmailSubject
            Dim noti_EmailCC As String = objEmail.Emailcc
            Dim noti_EmailBCC As String = objEmail.Emailbcc
            Dim NotificationType As String = "EMAIL"
            Dim noti_smtpserver As String = objEmail._smtpserver
            Dim noti_smtpuserid As String = objEmail._LoginID
            Dim noti_smtppassword As String = objEmail._LoginPassword
            Dim noti_smtpport As String = objEmail._smtpport
            Dim noti_ssl As Boolean = objEmail._SSL

            If ooonoti.notify(ConnectionStr, noti_MerchantID, NotificationType, noti_toNAme, noti_Message, noti_CCCode, noti_MobileNo, noti_EventCode, noti_toEmail, noti_FromNAme, noti_FromEmail, noti_EmailCC, noti_emailbcc, noti_EmailSubject, 3, noti_smtpserver, noti_smtpuserid, noti_smtppassword, noti_smtpport, noti_ssl) = False Then
                ErrorMSg = ooonoti.ErrorMSg
                Return False
            End If
            ooonoti = Nothing
            '***********************************************************

            Return True

        Catch ex As Exception
            ErrorMSg = ex.MEssage
            Return False
        End Try



    End Function


End Class


Public Class RuntimeUser
    Public ErrorMsg As String = ""
    Private connectionstring As String = System.Configuration.ConfigurationSettings.AppSettings("ConnStr")
    Public Function GetNotificationDetailsbyLoginID(ByVal LoginID As String, ByRef EmailAddress As String, ByRef MobileNo As String, Optional ByRef ToName As String = "") As Boolean
        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim dr As DataRow

        Try
            cn.open()
            cmd.CommandText = "Select * from secuserinfo where usr_loginid='" & LoginID & "'"
            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")
            For Each dr In ds.Tables("datarecords").Rows
                counter = counter + 1
                Dim lemail As String = ""
                If (dr("usr_email") & "").trim <> "" Then
                    lemail = (dr("usr_email") & "").trim
                End If
                'If (dr("usr_email2") & "").trim <> "" Then
                '    If lemail.trim <> "" Then
                '        lemail = lemail & ","
                '    End If
                '    lemail = lemail & (dr("usr_email2") & "").trim
                'End If
                'If (dr("usr_email3") & "").trim <> "" Then
                '    If lemail.trim <> "" Then
                '        lemail = lemail & ","
                '    End If
                '    lemail = lemail & (dr("usr_email3") & "").trim
                'End If

                EmailAddress = lemail  'dr("usr_email") & ""


                MobileNo = dr("usr_hpcc") & "" & dr("usr_hpno") & ""

                ToName = dr("usr_name") & ""
                Exit For
            Next
            cn.Close()
            cmd.Dispose()
            cn.Dispose()
            Return True
        Catch ex As Exception
            ErrorMsg = ex.Message
            Return False
        End Try



    End Function


End Class
