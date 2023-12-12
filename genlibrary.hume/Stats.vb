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

    Public Shared Sub trackstats(ByVal _p_type As String, Optional ByVal _p_status As String = "", Optional ByVal _p_conn As String = "")
        Dim lSQL As String

        'L - Login
		'E - Email Link

        Dim connectionstring As String = System.Configuration.ConfigurationSettings.AppSettings("ConnStr")
        lSQL = "Insert into sitestats(sss_loginid,sss_transdate,sss_merchantid,sss_filtercode,sss_isstaff,sss_custnum,sss_type,sss_portal,sss_useragent,sss_status,sss_adconnection) " & _
               "values ('" &weblib.loginuser & "',getdate(),'" & weblib.merchantid & "','" & weblib.filtercode & "'," & weblib.booleantobit(weblib.isstaff) & "," & _ 
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


End Class






