Imports Microsoft.VisualBasic
Imports System.Data.OleDB
Imports System.Data

Public Class clsWorkflow
    Public ErrorMsg As String = ""
    Public connectionstring As String = System.Configuration.ConfigurationSettings.AppSettings("ConnStr")

    Public Shared Function getLevelNameSQL(ByVal WorkFlowIDLinkField As String, ByVal WorkflowLevelField As String) As String
        Dim lsql As String = ""
        lsql = "Select top 1 tempwfi.wui_name from workflowitems tempwfi where tempwfi.wui_wid=" & WorkFlowIDLinkField & " and isnull(tempwfi.wui_no,0)=" & WorkflowLevelField


        '       lsql = "select top 1 tempwfi.wui_name from workflowitems tempwfi where tempwfi.wui_no in (select top " & WorkflowLevelField & " blaa.wui_no from workflowitems blaa where blaa.wui_wid = " & WorkFlowIDLinkField & " order by blaa.wui_no asc) and tempwfi.wui_wid = " & WorkFlowIDLinkField & " order by tempwfi.wui_no desc"


        Return lsql
    End Function
    Public Shared Function getUserNamebyWorkFlowSQL(ByVal WorkFlowIDLinkField As String, ByVal WorkflowLevelField As String) As String
        Dim lsql As String = ""
        lsql = "Select distinct usr_name from secuserinfo inner join wgrouprights on secuserinfo.usr_id = wur_uid and Charindex('''' + convert(varchar(max),wur_wgroupid) + '''',(Select top 1  '''' + replace(isnull(wui_rights,''),';;',''',''') + '0''' from workflowitems where wui_wid=" & WorkFlowIDLinkField & " and isnull(wui_no,0)=" & WorkflowLevelField & ")) <> 0"
        Return lsql
    End Function
    Public Shared Function getUserCodebyWorkFlowSQL(ByVal WorkFlowIDLinkField As String, ByVal WorkflowLevelField As String) As String
        Dim lsql As String = ""
        lsql = "Select distinct usr_code from secuserinfo inner join wgrouprights on secuserinfo.usr_id = wur_uid and Charindex('''' + convert(varchar(max),wur_wgroupid) + '''',(Select top 1  '''' + replace(isnull(wui_rights,''),';;',''',''') + '0''' from workflowitems where wui_wid=" & WorkFlowIDLinkField & " and isnull(wui_no,0)=" & WorkflowLevelField & ")) <> 0"
        Return lsql


    End Function
    Public Shared Function getLevelsbyWorkflowID(ByVal WorkFlowID As String) As String
        Dim lsql As String = ""
        '        lsql = "Select ROW_NUMBER() OVER(ORDER BY wui_no asc) AS Recno,workflowitems.* from workflowitems where workflowitems.wui_wid=" & WorkFlowID & " order by wui_no asc"

        lsql = "Select wui_no AS Recno,workflowitems.* from workflowitems where workflowitems.wui_wid=" & WorkFlowID & " order by wui_no asc"
        Return lsql
    End Function

    Public Function getAuditSQL(ByVal pLevel As String, ByVal pType As String, ByVal pRefNo As String, ByVal pDescription As String, ByVal pucode As String) As String

        If IsNumeric(pLevel) = False Then
            pLevel = "NULL"
        End If

        Return "Insert into workflowaudit (wfa_code,wfa_refno,wfa_description,wfa_ucode,wfa_createon,wfa_createby,wfa_merchantid,wfa_filtercode,wfa_level) Values " &
                        "('" & pType & "','" & pRefNo & "','" & pDescription & "','" & pucode & "',getdate(),'" & WebLib.LoginUser & "','" & WebLib.MerchantID & "','" & WebLib.FilterCode & "'," & pLevel & ")"

    End Function
    Public Function GetNextLevel(ByVal workflowid As String, ByVal pLevel As String) As String
        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim dr As DataRow
        Dim lLevel As String = ""
        If IsNumeric(pLevel) = False Then
            Return ""
            Exit Function

        End If
        If IsNumeric(workflowid) = False Then
            Return ""
            Exit Function
        End If

        Try
            cmd.CommandText = "Select top 1 wui_no from workflowitems where wui_wid=" & workflowid & " and isnull(wui_no,0)>" & pLevel & " order by wui_no asc"
            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")
            For Each dr In ds.Tables("datarecords").Rows
                lLevel = dr("wui_no") & ""
                Exit For
            Next
            cn.Close()
            cmd.Dispose()
            cn.Dispose()

            Return lLevel

        Catch ex As Exception
            ErrorMsg = ex.Message
            Return ""
        End Try
    End Function
    Public Function GetPreviousLevel(ByVal workflowid As String, ByVal pLevel As String) As String
        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim dr As DataRow
        Dim lLevel As String = ""
        If IsNumeric(pLevel) = False Then
            Return ""
            Exit Function

        End If
        If IsNumeric(workflowid) = False Then
            Return ""
            Exit Function
        End If

        Try
            cmd.CommandText = "Select top 1 wui_no from workflowitems where wui_wid=" & workflowid & " and isnull(wui_no,0)<" & pLevel & " order by wui_no desc"
            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")
            For Each dr In ds.Tables("datarecords").Rows
                lLevel = dr("wui_no") & ""
                Exit For
            Next
            cn.Close()
            cmd.Dispose()
            cn.Dispose()

            Return lLevel

        Catch ex As Exception
            ErrorMsg = ex.Message
            Return ""
        End Try
    End Function
    Public Function GetLevelNobySequence(ByVal workflowid As String, ByVal pLevel As String) As String
        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim dr As DataRow
        Dim lLevel As String = ""
        If IsNumeric(pLevel) = False Then
            Return ""
            Exit Function

        End If
        If IsNumeric(workflowid) = False Then
            Return ""
            Exit Function
        End If

        Try
            cmd.CommandText = "Select ROW_NUMBER() OVER(ORDER BY wui_no asc) AS Recno from workflowitems where wui_wid=" & workflowid & " and isnull(wui_no,0)<=" & pLevel & " order by wui_no desc"
            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")
            For Each dr In ds.Tables("datarecords").Rows
                lLevel = dr("Recno") & ""
                Exit For
            Next
            cn.Close()
            cmd.Dispose()
            cn.Dispose()

            Return lLevel

        Catch ex As Exception
            ErrorMsg = ex.Message
            Return ""
        End Try
    End Function


End Class
Public Class clsWorkflowEmail
    Public ErrorMsg As String = ""
    Public connectionstring As String = System.Configuration.ConfigurationSettings.AppSettings("ConnStr")
    Public ReadOnly Property GetWorflowCustomFields()
        Get
            Return "#WorkFlowName#|#SecureURL#|#DocumentNo#|#CustomerName#|#CompanyName#"
        End Get
    End Property
    Private Function getWorflowEmailSQL(ByVal WorkFlowIDLinkField As String, ByVal WorkflowLevelField As String, ByVal WorkFlowAction As String, Optional ByVal Version2 As Boolean = False) As String
        Dim lfieldname As String = ""
        Select Case WorkFlowAction

            Case "A"
                lfieldname = "wui_emailA"
            Case "U"
                If Version2 = True Then
                    lfieldname = "wui_emailS"
                Else
                    lfieldname = "wui_emailA"
                End If
            Case "R"
                lfieldname = "wui_emailR"
            Case "C"
                lfieldname = "wui_emailC"
            Case "G"
                lfieldname = "wui_rights"
            Case "N"
                lfieldname = "wui_emailN"
            Case "L", "S"
                lfieldname = "convert(varchar,wui_seq)" ' need to be empty, not allow to send to customer

            Case Else
                Return ""
                Exit Function
        End Select

        Dim lsql As String = ""
        'lsql = "select * from(Select distinct ROW_NUMBER() OVER (ORDER BY usr_loginid) as rownumber, usr_loginid,usr_name,usr_email,wur_wgroupid from secuserinfo inner join wgrouprights on secuserinfo.usr_id = wur_uid and Charindex('''' + convert(varchar(max),wur_wgroupid) + '''',(Select top 1  '''' + replace(isnull(" & lfieldname & ",''),';;',''',''') + '0''' from workflowitems where wui_wid=" & WorkFlowIDLinkField & " and isnull(wui_no,0)=" & WorkflowLevelField & ")) <> 0  where RTRIM(ISNULL(usr_email,'')) <> '') a "
        ' do not trigger email to disabled user 
        lsql = "select * from(Select distinct ROW_NUMBER() OVER (ORDER BY usr_loginid) as rownumber, usr_loginid,usr_name,usr_email,wur_wgroupid from secuserinfo inner join wgrouprights on secuserinfo.usr_id = wur_uid And Charindex('''' + convert(varchar(max),wur_wgroupid) + '''',(Select top 1  '''' + replace(isnull(" & lfieldname & ",''),';;',''',''') + '0''' from workflowitems where wui_wid=" & WorkFlowIDLinkField & " And isnull(wui_no,0)=" & WorkflowLevelField & ")) <> 0  where RTRIM(ISNULL(usr_email,'')) <> '' and usr_disable = 0 ) temptable "

        Return lsql

    End Function
    Public Function NotifyUsers(ByVal WorkFlowID As String, ByVal WorkflowLevelField As String, ByVal WorkFlowAction As String, ByVal docnamespace As String, ByVal workflowuniqueid As String, Optional ByVal Version2 As Boolean = False, Optional ByVal adhocSendToCust As String = "", Optional ByVal adhocSendToEmail As String = "") As Boolean
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

        Select Case WorkFlowAction
            Case "U"
                lTemplate = "wflowrou"
            Case "A"
                lTemplate = "wflowapp"
            Case "R"
                lTemplate = "wflowrej"
            Case "C"
                lTemplate = "wflowcan"

            Case "N"
                lTemplate = "wflownot"
            Case "L"
                lTemplate = "wflowclo"
            Case "S"
                lTemplate = "wflowsub"
            Case Else
                Return ""
                Exit Function
        End Select

        If objEmail.GetEmailTemplate(lTemplate, "GENERAL") = True Then

            orgSubject = objEmail.EmailSubject
            orgEmail = objEmail.EmailBody

        Else
            ErrorMsg = "Error Init Sending"
            Return False
            Exit Function
        End If


        If (orgEmail.Contains("#CustomerName#") Or orgSubject.Contains("#CustomerName#")) Or (orgEmail.Contains("#WorkFlowName#") Or orgSubject.Contains("#WorkFlowName#")) Or (orgEmail.Contains("#DocumentNo#") Or orgSubject.Contains("#DocumentNo#")) Or (orgEmail.Contains("#CompanyName#") Or orgSubject.Contains("#CompanyName#")) Then
            Try

                Call EmailFieldsReplace(workflowuniqueid, orgSubject, orgEmail, adhocSendToCust)
            Catch ex As Exception

            End Try

        End If



        Try
            cmd.CommandText = getWorflowEmailSQL(WorkFlowID, WorkflowLevelField, WorkFlowAction, Version2)

            If (cmd.commandtext & "").trim = "" Then
                ErrorMsg = "Unable to retrieve query"
                Return False
                Exit Function
            End If
            cn.open()
            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")

            '************************ filterring 

            Dim getmodule As String = WebLib.GetValue("workflowstatus", "wst_module", "wst_ucode", "'" & workflowuniqueid & "'", "", "")
            Dim gettablename As String = ""
            Dim getcustomername As String = ""
            Dim filterstr As String = ""

            If getmodule = "zcustom_dn" Or getmodule = "zcustom_dn2" Or getmodule = "zcustom_dn3" Or getmodule = "zcustom_cn" Or getmodule = "zcustom_cn2" Or getmodule = "zcustom_cn3" Then
                gettablename = "zcustom_dncn"
            ElseIf getmodule = "zcustom_ccr" Or getmodule = "zcustom_ccrp" Or getmodule = "zcustom_ccrs" Or getmodule = "zcustom_ccrc" Then
                gettablename = "zcustom_ccr"
            ElseIf getmodule = "zcustom_clexceed" Then
                gettablename = "zcustom_climit"
            Else
                gettablename = getmodule
            End If

            Dim getaccountholderId As String = WebLib.GetValue(gettablename, "cus_accountholder", "cus_ucode", "'" & workflowuniqueid & "'", "", "")

            If gettablename = "zcustom_ccr" Then
                getcustomername = WebLib.GetValue(gettablename, "cus_distributor", "cus_ucode", "'" & workflowuniqueid & "'", "", "")
            Else
                getcustomername = WebLib.GetValue(gettablename, "cus_company", "cus_ucode", "'" & workflowuniqueid & "'", "", "")
            End If

            Dim dt As DataTable = ds.Tables("datarecords")

            Dim dv As DataView = New System.Data.DataView(dt)

            '' wur_wgroupid='49' - Customer Account Holder
            Dim foundRow() As DataRow = dt.Select("wur_wgroupid='49' and usr_loginid <> '" & WebLib.GetValue("secuserinfo", "usr_loginid", "usr_code", "'" & getaccountholderId & "'", "", "") & "' ")
            '' wur_wgroupid='48' - Customer 
            Dim foundRow1() As DataRow = dt.Select("wur_wgroupid='48' and usr_loginid <> '" & WebLib.GetValue("secuserinfo", "usr_loginid", "usr_name", "'" & getcustomername & "'", "", "") & "' ")

            For Each item As DataRow In foundRow
                filterstr = filterstr & " and rownumber <> '" & item("rownumber") & "' "
            Next
            For Each item As DataRow In foundRow1
                filterstr = filterstr & " and rownumber <> '" & item("rownumber") & "' "
            Next

            If filterstr <> "" Then
                dv.RowFilter = filterstr.Substring(4)
            End If
            dt = dv.ToTable

            '************************ 

            If dt.Rows.Count = 0 And (WorkFlowAction = "L" Or WorkFlowAction = "S") And adhocSendToCust = "Y" Then
                Dim nrow As DataRow = dt.NewRow
                nrow(0) = 0
                nrow(1) = ""
                nrow(2) = ""
                nrow(3) = adhocSendToEmail
                nrow(4) = 0
                dt.Rows.Add(nrow)
            End If



            'For Each dr In ds.Tables("datarecords").Rows
            For Each dr In dt.Rows


                counter = counter + 1


                objEmail.EmailSubject = orgSubject
                objEmail.EmailBody = orgEmail


                '#CustomerName#
                '#WorkFlowName#
                '#DocumentNo#



                If orgEmail.Contains("#SecureURL#") Or orgSubject.Contains("#SecureURL#") Then
                    Dim oore As New Vifeandi.Custom.Redirect
                    Dim oowl As New WebLib
                    Dim lcode As String = ""
                    Dim lURL As String
                    lcode = oore.web_redirect("WORKFLOW", 30, 10, dr("usr_loginid") & "", WorkFlowID, docnamespace, workflowuniqueid)
                    If (lcode & "").Trim <> "" Then
                        lURL = oowl.AbsoluteWebPath & "/modules/general/redirect.aspx?c=" & lcode
                        objEmail.EmailBody = objEmail.EmailBody.replace("#SecureURL#", lURL)
                        objEmail.EmailSubject = objEmail.EmailSubject.replace("#SecureURL#", lURL)

                    End If
                    oore = Nothing
                    oowl = Nothing
                End If



                objEmail.Emailto = dr("usr_email") & ""
                objEmail.FromName = "Workflow Notification"



                '************ NOTIFICATION CENTER ********************
                Dim ooonoti As New ViFeandi.NotificationCenter

                Dim ConnectionStr As String = connectionstring   '"Provider=SQLOLEDB;Data Source=HC-EPI-APS1;Initial Catalog=CustomerPortalNotification;User ID=USR_PORTAL;Password=1qazXSW@"
                Dim noti_MerchantID As String = Weblib.MerchantID
                Dim noti_Message As String = objEmail.EmailBody
                Dim noti_CCCode As String = ""
                Dim noti_MobileNo As String = ""
                Dim noti_EventCode As String = "WORKFLOW"
                Dim noti_toNAme As String = objEmail.EmailTo
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

                If ooonoti.notify(ConnectionStr, noti_MerchantID, NotificationType, noti_toNAme, noti_Message, noti_CCCode, noti_MobileNo, noti_EventCode, noti_toEmail, noti_FromNAme, noti_FromEmail, noti_EmailCC, noti_EmailBCC, noti_EmailSubject, 10, noti_smtpserver, noti_smtpuserid, noti_smtppassword, noti_smtpport, noti_ssl) = False Then
                    '                    ooonoti = Nothing
                    '                   ErrorMsg = ooonoti.ErrorMsg
                    '                  Return False
                    '                 Exit Function
                End If

                ooonoti = Nothing
                '***********************************************************


                '    If objEmail.sendEmail() = False Then
                'ErrorMsg = Weblib.ErrorTrap
                'Return False
                'Exit Function
                'End If



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
    Private Function EmailFieldsReplace(ByVal workflow_ucode As String, ByRef EmailSubject As String, ByRef Emailbody As String, Optional ByVal adhocSendToCust As String = "") As Boolean
        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim dr As DataRow

        Try
            cn.open()
            cmd.CommandText = "Select wst_param1,wst_param2,wst_param3,wst_param4,wst_subject, (select br_name from branch where br_code ='" & WebLib.LoginUserCompanySelected & "') as companyname from workflowstatus where wst_ucode='" & workflow_ucode & "'"
            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")
            For Each dr In ds.Tables("datarecords").Rows


                EmailSubject = EmailSubject.Replace("#CustomerName#", dr("wst_param1") & "")
                EmailSubject = EmailSubject.Replace("#WorkFlowName#", dr("wst_subject") & "")
                If adhocSendToCust = "Y" Then
                    EmailSubject = EmailSubject.Replace("#DocumentNo#", dr("wst_param3") & "")
                Else
                    EmailSubject = EmailSubject.Replace("#DocumentNo#", dr("wst_param2") & "")
                End If
                EmailSubject = EmailSubject.Replace("#CompanyName#", dr("companyname") & "")

                '********************************************************************************************
                Emailbody = Emailbody.Replace("#CustomerName#", dr("wst_param1") & "")
                Emailbody = Emailbody.Replace("#WorkFlowName#", dr("wst_subject") & "")
                If adhocSendToCust = "Y" Then
                    Emailbody = Emailbody.Replace("#DocumentNo#", dr("wst_param3") & "")
                Else
                    Emailbody = Emailbody.Replace("#DocumentNo#", dr("wst_param2") & "")
                End If
                Emailbody = Emailbody.Replace("#CompanyName#", dr("companyname") & "")

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