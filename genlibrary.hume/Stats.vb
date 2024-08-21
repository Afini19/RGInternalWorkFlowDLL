Imports System
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Data.Odbc

Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Drawing.Drawing2D
Imports System.Drawing.Text
Imports System.Configuration
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.HtmlControls
Imports System.Text
Imports System.Collections.Generic
Imports System.Xml
Imports System.Security.Cryptography
Imports System.Net.Mail
Imports System.Text.RegularExpressions

Public Class WebStats

    Public pFieldNames As String = ""
    Public pJoinFields As String = ""
    Public TableName As String = ""
    Public _searchfilter As String = ""

    Dim connectionstring As String = System.Configuration.ConfigurationSettings.AppSettings("ConnStr")

    Public Shared Sub trackstats(ByVal _p_type As String, Optional ByVal _p_status As String = "", Optional ByVal _p_conn As String = "")
        Dim lSQL As String

        'L - Login
        'E - Email Link

        Dim connectionstring As String = System.Configuration.ConfigurationSettings.AppSettings("ConnStr")
        lSQL = "Insert into sitestats(sss_loginid,sss_transdate,sss_merchantid,sss_filtercode,sss_isstaff,sss_custnum,sss_type,sss_portal,sss_useragent,sss_status,sss_adconnection) " &
               "values ('" & weblib.loginuser & "',getdate(),'" & weblib.merchantid & "','" & weblib.filtercode & "'," & weblib.booleantobit(weblib.isstaff) & "," &
                    "'" & weblib.custnum & "','" & _p_type & "',NULL,'" & HttpContext.Current.Request.UserAgent & "', '" & _p_status & "','" & _p_conn & "')"


        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()

        cn.Open()
        cmd.CommandText = lSQL
        cmd.Connection = cn
        cmd.ExecuteNonQuery()

        cn.Close()
        cmd.Dispose()
        cn.Dispose()

    End Sub

    Public Function GetPendingCRCStats() As String
        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDb.OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim stats As String = ""

        TableName = "zcustom_crC"

        pFieldNames = "( Select workflowstatus.*, " & TableName & ".*,ApprovalLevelName = wui_name, tempwfi.[value] as [rights] "

        pJoinFields = "inner join workflowstatus on cus_ucode = wst_ucode " &
                        "left join (Select * From workflowitems Cross apply string_split(replace(isnull(wui_rights,''),';;',',') + '0',',') where value <> 0 ) tempwfi on tempwfi.wui_wid=wst_workflowid and isnull(tempwfi.wui_no,0)=wst_level " &
                        "left outer join secuserinfo on cus_createby = usr_loginid ) as k "
        '"inner join wgrouprights on wur_wgroupid = rights and wur_wgroupid not in (49) left join secuserinfo secwur on wur_uid = secwur.usr_id "

        _searchfilter = " wst_status='Pending' and (cus_createby='" & WebLib.LoginUser & "' or '" & WebLib.LoginUser & "' in (" & clsWorkflow.getUserCodebyWorkFlowSQL("wst_workflowid", "wst_level") & ")) "

        Try

            cn.Open()
            cmd.CommandText = "Select count(wst_id) as count from " + " " + pFieldNames + " from " + TableName + " " + pJoinFields + " where " + _searchfilter
            LogtheAudit(cmd.CommandText)
            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")
            cn.Close()
            cmd.Dispose()
            cn.Dispose()

            For Each dr In ds.Tables(0).Rows
                stats = dr("count").ToString()
            Next

            Return stats

        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Function GetPendingCRIStats() As String
        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDb.OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim stats As String = ""

        TableName = "zcustom_crI"

        pFieldNames = "( Select workflowstatus.*, " & TableName & ".*,ApprovalLevelName = wui_name, tempwfi.[value] as [rights] "

        pJoinFields = "inner join workflowstatus on cus_ucode = wst_ucode " &
                        "left join (Select * From workflowitems Cross apply string_split(replace(isnull(wui_rights,''),';;',',') + '0',',') where value <> 0 ) tempwfi on tempwfi.wui_wid=wst_workflowid and isnull(tempwfi.wui_no,0)=wst_level " &
                        "left outer join secuserinfo on cus_createby = usr_loginid ) as k "
        '"inner join wgrouprights on wur_wgroupid = rights and wur_wgroupid not in (49) left join secuserinfo secwur on wur_uid = secwur.usr_id "

        _searchfilter = " wst_status='Pending' and (cus_createby='" & WebLib.LoginUser & "' or '" & WebLib.LoginUser & "' in (" & clsWorkflow.getUserCodebyWorkFlowSQL("wst_workflowid", "wst_level") & ")) "

        Try

            cn.Open()
            cmd.CommandText = "Select count(wst_id) as count from " + " " + pFieldNames + " from " + TableName + " " + pJoinFields + " where " + _searchfilter
            LogtheAudit(cmd.CommandText)
            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")
            cn.Close()
            cmd.Dispose()
            cn.Dispose()

            For Each dr In ds.Tables(0).Rows
                stats = dr("count").ToString()
            Next

            Return stats

        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Function GetPendingCRSStats() As String
        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDb.OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim stats As String = ""

        TableName = "zcustom_crS"

        pFieldNames = "( Select workflowstatus.*, " & TableName & ".*,ApprovalLevelName = wui_name, tempwfi.[value] as [rights] "

        pJoinFields = "inner join workflowstatus on cus_ucode = wst_ucode " &
                        "left join (Select * From workflowitems Cross apply string_split(replace(isnull(wui_rights,''),';;',',') + '0',',') where value <> 0 ) tempwfi on tempwfi.wui_wid=wst_workflowid and isnull(tempwfi.wui_no,0)=wst_level " &
                        "left outer join secuserinfo on cus_createby = usr_loginid ) as k "
        '"inner join wgrouprights on wur_wgroupid = rights and wur_wgroupid not in (49) left join secuserinfo secwur on wur_uid = secwur.usr_id "

        _searchfilter = " wst_status='Pending' and (cus_createby='" & WebLib.LoginUser & "' or '" & WebLib.LoginUser & "' in (" & clsWorkflow.getUserCodebyWorkFlowSQL("wst_workflowid", "wst_level") & ")) "

        Try

            cn.Open()
            cmd.CommandText = "Select count(wst_id) as count from " + " " + pFieldNames + " from " + TableName + " " + pJoinFields + " where " + _searchfilter
            LogtheAudit(cmd.CommandText)
            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")
            cn.Close()
            cmd.Dispose()
            cn.Dispose()

            For Each dr In ds.Tables(0).Rows
                stats = dr("count").ToString()
            Next

            Return stats

        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Function GetPendingSTIStats() As String
        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDb.OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim stats As String = ""

        TableName = "zcustom_stI"

        pFieldNames = "( Select workflowstatus.*, " & TableName & ".*,ApprovalLevelName = wui_name, tempwfi.[value] as [rights] "

        pJoinFields = "inner join workflowstatus on cus_ucode = wst_ucode " &
                        "left join (Select * From workflowitems Cross apply string_split(replace(isnull(wui_rights,''),';;',',') + '0',',') where value <> 0 ) tempwfi on tempwfi.wui_wid=wst_workflowid and isnull(tempwfi.wui_no,0)=wst_level " &
                        "left outer join secuserinfo on cus_createby = usr_loginid ) as k "
        '"inner join wgrouprights on wur_wgroupid = rights and wur_wgroupid not in (49) left join secuserinfo secwur on wur_uid = secwur.usr_id "

        _searchfilter = " wst_status='Pending' and (cus_createby='" & WebLib.LoginUser & "' or '" & WebLib.LoginUser & "' in (" & clsWorkflow.getUserCodebyWorkFlowSQL("wst_workflowid", "wst_level") & ")) "

        Try

            cn.Open()
            cmd.CommandText = "Select count(wst_id) as count from " + " " + pFieldNames + " from " + TableName + " " + pJoinFields + " where " + _searchfilter
            LogtheAudit(cmd.CommandText)
            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")
            cn.Close()
            cmd.Dispose()
            cn.Dispose()

            For Each dr In ds.Tables(0).Rows
                stats = dr("count").ToString()
            Next

            Return stats

        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Function GetPendingSTSStats() As String
        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDb.OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim stats As String = ""

        TableName = "zcustom_stS"

        pFieldNames = "( Select workflowstatus.*, " & TableName & ".*,ApprovalLevelName = wui_name, tempwfi.[value] as [rights] "

        pJoinFields = "inner join workflowstatus on cus_ucode = wst_ucode " &
                        "left join (Select * From workflowitems Cross apply string_split(replace(isnull(wui_rights,''),';;',',') + '0',',') where value <> 0 ) tempwfi on tempwfi.wui_wid=wst_workflowid and isnull(tempwfi.wui_no,0)=wst_level " &
                        "left outer join secuserinfo on cus_createby = usr_loginid ) as k "
        '"inner join wgrouprights on wur_wgroupid = rights and wur_wgroupid not in (49) left join secuserinfo secwur on wur_uid = secwur.usr_id "

        _searchfilter = " wst_status='Pending' and (cus_createby='" & WebLib.LoginUser & "' or '" & WebLib.LoginUser & "' in (" & clsWorkflow.getUserCodebyWorkFlowSQL("wst_workflowid", "wst_level") & ")) "

        Try

            cn.Open()
            cmd.CommandText = "Select count(wst_id) as count from " + " " + pFieldNames + " from " + TableName + " " + pJoinFields + " where " + _searchfilter
            LogtheAudit(cmd.CommandText)
            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")
            cn.Close()
            cmd.Dispose()
            cn.Dispose()

            For Each dr In ds.Tables(0).Rows
                stats = dr("count").ToString()
            Next

            Return stats

        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Sub LogtheAudit(ByVal theMessage As String)
        Dim strFile As String = "c:\officeonelog\ErrorLogWF.txt"
        Dim fileExists As Boolean = File.Exists(strFile)

        Try

            Using sw As New StreamWriter(File.Open(strFile, FileMode.Append))
                sw.WriteLine(DateTime.Now & " - " & theMessage)
            End Using
        Catch ex As Exception

        End Try
    End Sub
End Class






