Imports Microsoft.VisualBasic
Imports System.Web
Imports System.IO
Public Class Utils
    Public Shared Function checktext(ByVal pStr As String) As String
        'Return pstr
	Return Replace(pStr, "'", "''")
    End Function

    Public Shared Function SendEmailbyINI(ByVal ConnectionString As String, ByVal TransType As String, ByVal SettingsIniFilePath As String, ByVal EmailFormatFileURL As String, ByVal EmailTo As String, ByVal AdditionalParam1 As String, ByVal AdditionalParam2 As String, ByVal AdditionalParam3 As String, ByVal AdditionalParam4 As String, ByVal AdditionalParam5 As String, Optional ByVal EmailCC As String = "", Optional ByVal EmailBCC As String = "") As Boolean

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
        If SMTPSSLDATA.ToLower = "true" Then
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

            Dim ConnectionStr As String = ConnectionString
            Dim noti_MerchantID As String = WebLib.MerchantID
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

            If ooonoti.notify(ConnectionStr, noti_MerchantID, NotificationType, noti_toNAme, noti_Message, noti_CCCode, noti_MobileNo, noti_EventCode, noti_toEmail, noti_FromNAme, noti_FromEmail, noti_EmailCC, noti_EmailBCC, noti_EmailSubject, 3, noti_smtpserver, noti_smtpuserid, noti_smtppassword, noti_smtpport, noti_ssl) = False Then


            End If

            ooonoti = Nothing

            Return True
        Catch ex As Exception
            Return False
        End Try
        ooini = Nothing
        sendmail = Nothing



    End Function

End Class
