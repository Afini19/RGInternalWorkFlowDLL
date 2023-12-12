Imports Microsoft.VisualBasic
Imports System.Data.OleDB
Imports System.Data
Imports System.IO
Imports System.Web
Public Class mobileapproval
    Private Sub logdata(ByVal logdata As String)
        Dim SR As StreamWriter
        Dim Log_filename As String = "log_mobileapproval.txt"
        Try

            If Not File.Exists(System.AppDomain.CurrentDomain.BaseDirectory() & Log_filename) Then
                SR = File.CreateText(System.AppDomain.CurrentDomain.BaseDirectory() & Log_filename)
            Else
                SR = New StreamWriter(System.AppDomain.CurrentDomain.BaseDirectory() & Log_filename, True)
            End If

            SR.WriteLine(Format(DateTime.Now, "dd/MM/yyyy hh:mm:ss") & " : " & logdata)
            SR.Close()
        Catch ex As Exception

        End Try

    End Sub
    Public Function ActivateApproval(ByVal ConnectionString As String, ByVal SettingsIniFilePath As String, ByVal RedirectURL As String, ByVal EmailFormatFileURL As String, ByVal UniqueReference As String, ByVal TransType As String, ByVal EnableApproval As Boolean, ByVal EnableReject As Boolean, ByVal EnableCancellation As Boolean, ByVal EmailTo As String, ByVal Emailcc As String, ByVal Emailbcc As String, ByVal AdditionalParam1 As String, ByVal AdditionalParam2 As String, ByVal AdditionalParam3 As String, ByVal AdditionalParam4 As String, ByVal AdditionalParam5 As String, ByVal EmailParam1 As String, ByVal EmailParam2 As String) As String

        Dim rooturl As String = ""

        Try
            rooturl = RedirectURL               ' oowl.AbsoluteWebPath & "/gps/redirect.aspx?c=" & lcode

            Dim Parameter1 As String = EmailParam1
            Dim PArameter2 As String = EmailParam2
            Dim Parameter3 As String = ""   'Used as approval link
            Dim Parameter4 As String = ""   'Used as reject link
            Dim Parameter5 As String = ""   'Used as cancel link
            Dim ltranscode As String = ""

            Dim oore As New Vifeandi.Custom.Redirect
            Dim oo As New ViFeandi.General
            ltranscode = oo.getUniqueCode(4)
            oo = Nothing

            If (RedirectURL & "").Trim = "" Then
                logdata("Invalid Redirect URL")
                Return ""
                Exit Function
            End If
            If (SettingsIniFilePath & "").Trim = "" Then
                logdata("INI File Path Error")
                Return ""
                Exit Function
            End If
            If (ConnectionString & "").Trim = "" Then
                logdata("Invalid DB Connection String")
                Return ""
                Exit Function
            End If

            If (EmailFormatFileURL & "").Trim = "" Then
                logdata("Invalid Email Format File URL")
                Return ""
                Exit Function
            End If

            If EnableApproval = True Then
                Dim lcode As String = ""
                wait(1)
                lcode = oore.web_redirect("APPROVE", 30, 3, "", "", ltranscode, UniqueReference)
                If (lcode & "").Trim = "" Then
                    logdata("Error Getting Unique Redirect Code")
                    Return ""
                    Exit Function
                End If
                Parameter3 = rooturl & "redirect.aspx?c=" & lcode
            End If

            If EnableReject = True Then
                Dim lcode As String = ""
                wait(1)
                lcode = oore.web_redirect("REJECT", 30, 3, "", "", ltranscode, UniqueReference)
                If (lcode & "").Trim = "" Then
                    logdata("Error Getting Unique Redirect Code")
                    Return ""
                    Exit Function
                End If
                Parameter4 = rooturl & "redirect.aspx?c=" & lcode
            End If

            If EnableCancellation = True Then
                Dim lcode As String = ""
                wait(1)
                lcode = oore.web_redirect("CANCEL", 30, 3, "", "", ltranscode, UniqueReference)
                If (lcode & "").Trim = "" Then
                    logdata("Error Getting Unique Redirect Code")
                    Return ""
                    Exit Function
                End If
                Parameter5 = rooturl & "redirect.aspx?c=" & lcode
            End If


            oore = Nothing

            Dim lsql As String = ""
            Dim ltranstype As String = Transtype
            Dim lstatus As String = "PENDING"

            lsql = "Insert into ApprovalStatus (as_uniqueno,as_refno,as_transtype,as_createdt,as_param1,as_param2,as_param3,as_param4,as_param5,as_status) Values ("
            lsql = lsql & "'" & ltranscode & "','" & UniqueReference & "','" & ltranstype & "',getdate(),'" & AdditionalParam1 & "','" & AdditionalParam2 & "','" & AdditionalParam3 & "','" & AdditionalParam4 & "','" & AdditionalParam5 & "','" & lstatus & "')"

            If SQLCommand_EXEC(ConnectionString, lsql) <> "Y" Then
                Call logdata("Error Saving : " & UniqueReference)
                Return ""
                Return False
            End If

            If sendemail(ConnectionString, SettingsIniFilePath, TransType, EmailFormatFileURL, EmailTo, Parameter1, PArameter2, Parameter3, Parameter4, Parameter5, emailcc, EmailBCC) = False Then
                Call logdata("Error Sending Mail : " & UniqueReference)
                Return ""
                Return False
            End If
            '**************************************************

            Return ltranscode

        Catch ex As Exception
            Call logdata("Activating Approval : " & ex.message)


        End Try

    End Function
    Private Function SQLCommand_EXEC(ByVal _p_connstr As String, ByVal _p_sql As String) As String
        Dim strConnection As String = ""
        Dim dr As DataRow
        Dim cnHQ As New OleDb.OleDbConnection
        Dim cmdHQ As New OleDb.OleDbCommand
        Dim trans As OleDb.OleDbTransaction

        strConnection = _p_connstr

        Try
            cnHQ = New OleDb.OleDbConnection(strConnection)


            cnHQ.Open()
            cmdHQ.Connection = cnHQ

            trans = cnHQ.BeginTransaction()
            cmdHQ.Transaction = trans

            cmdHQ.CommandText = _p_sql
            cmdHQ.ExecuteNonQuery()

            trans.Commit()

            Return "Y"

            cmdHQ.Dispose()
            cnHQ.Close()
            cnHQ.Dispose()
        Catch ex As Exception

            trans.Rollback()
            Return "N"
            Call logdata(ex.Message)
        Finally

        End Try
    End Function
    Private Function sendemail(ByVal ConnectionString As String, ByVal SettingsIniFilePath As String, ByVal TransType As String, ByVal EmailFormatFileURL As String, ByVal EmailTo As String, ByVal AdditionalParam1 As String, ByVal AdditionalParam2 As String, ByVal AdditionalParam3 As String, ByVal AdditionalParam4 As String, ByVal AdditionalParam5 As String, Optional ByVal EmailCC As String = "", Optional ByVal EmailBCC As String = "") As Boolean

        Dim lemail As String = ""
        Dim bodyContent As String = ""
        Dim sendmail As New ViFeandiEmail.Email2u
        Dim inifilepath As String = SettingsIniFilePath
        Dim ooini As New ViFeandi.VIFEANDI_APP_INI

        Dim lFromName As String = ooini.INIRead(inifilepath, "emailroute", "EmailFromName", "")
        Dim lSubject As String = ooini.INIRead(inifilepath, "emailroute", "EmailSubject", "")
        Dim SMTPLoginID As String = ooini.INIRead(inifilepath, "smtpserver", "smtploginid", "")
        Dim SMTPLoginPwd As String = ooini.INIRead(inifilepath, "smtpserver", "smtppwd", "")
        Dim SMTPFROMEMAIL As String = ooini.INIRead(inifilepath, "emailroute", "EmailFrom", "")
        Dim SMTPSERVER As String = ooini.INIRead(inifilepath, "smtpserver", "smtpserver", "")
        Dim SMTPPORt As String = ooini.INIRead(inifilepath, "smtpserver", "smtpport", "")
        Dim SMTPSSLDATA As String = ooini.INIRead(inifilepath, "smtpserver", "smtpssl", "")
        Dim SMTPSSL As Boolean = False

        If SMTPSSLDATA.tolower = "true" Then
            SMTPSSL = True
        End If

        Dim EmailToEmail As String = EmailTo
        Dim EmailToName As String = EmailTo

        bodyContent = File.ReadAllText(HttpContext.Current.Server.MapPath(EmailFormatFileURL))

        bodyContent = bodyContent.Replace("##Parameter1##", AdditionalParam1)
        bodyContent = bodyContent.Replace("##Parameter2##", AdditionalParam2)
        bodyContent = bodyContent.Replace("##Parameter3##", AdditionalParam3)
        bodyContent = bodyContent.Replace("##Parameter4##", AdditionalParam4)
        bodyContent = bodyContent.Replace("##Parameter4##", AdditionalParam5)

        lSubject = lSubject.Replace("##Parameter1##", AdditionalParam1)
        lSubject = lSubject.Replace("##Parameter2##", AdditionalParam2)
        lSubject = lSubject.Replace("##Parameter3##", AdditionalParam3)
        lSubject = lSubject.Replace("##Parameter4##", AdditionalParam4)
        lSubject = lSubject.Replace("##Parameter5##", AdditionalParam5)


        Try

            '************ NOTIFICATION CENTER ********************
            Dim ooonoti As New ViFeandi.NotificationCenter

            Dim ConnectionStr As String = connectionstring
            Dim noti_MerchantID As String = Weblib.MerchantID
            Dim noti_Message As String = bodyContent
            Dim noti_CCCode As String = ""
            Dim noti_MobileNo As String = ""
            Dim noti_EventCode As String = TransType
            Dim noti_toNAme As String = EmailToName
            Dim noti_toEmail As String = EmailToName
            Dim noti_FromNAme As String = lFromName
            Dim noti_FromEmail As String = SMTPFROMEMAIL
            Dim noti_EmailSubject As String = lSubject
            Dim noti_EmailCC As String = EmailCC
            Dim noti_EmailBCC As String = EmailBCC
            Dim NotificationType As String = "EMAIL"
            Dim noti_smtpserver As String = SMTPSERVER
            Dim noti_smtpuserid As String = SMTPLoginID
            Dim noti_smtppassword As String = SMTPLoginPwd
            Dim noti_smtpport As String = SMTPPORt
            Dim noti_ssl As Boolean = SMTPSSL

            If ooonoti.notify(ConnectionStr, noti_MerchantID, NotificationType, noti_toNAme, noti_Message, noti_CCCode, noti_MobileNo, noti_EventCode, noti_toEmail, noti_FromNAme, noti_FromEmail, noti_EmailCC, noti_emailbcc, noti_EmailSubject, 3, noti_smtpserver, noti_smtpuserid, noti_smtppassword, noti_smtpport, noti_ssl) = False Then


            End If

            ooonoti = Nothing
            '***********************************************************
            '            sendmail.InitSMTPServer(SMTPSERVER, SMTPFROMEMAIL, SMTPPORt, SMTPLoginID, SMTPLoginPwd)
            '           sendmail.SendEmail(EmailToEmail, SMTPFROMEMAIL, EmailToName, lFromName, lSubject, bodyContent, True, False, EmailCC, EmailBCC)

            Return True
        Catch ex As Exception
            logdata("SEND MAIL ERROR : " & ex.Message)
            Return False
        End Try
        ooini = Nothing
        sendmail = Nothing



    End Function
    Private Sub wait(ByVal seconds As Integer)
        System.Threading.Thread.Sleep(100)
    End Sub
End Class