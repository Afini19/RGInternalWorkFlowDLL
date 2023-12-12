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

public class WebLib
    Public NoPermissionToDelete As String = "You do not have permisson to delete record."
    Public RecordNotFound As String = "Record does not exist."
    Public DeletedSuccessfullyMessage As String = "Record deleted successfully."
    Public DeleteFailedMessage As String = "Failed to delete the specified record."
    Public UpdatedSuccessfullyMessage As String = "Record updated successfully."
    Public UpdateFailedMessage As String = "Failed to update the specified record."
    Public AddedSuccessfullyMessage As String = "Record added successfully."
    Public AddFailedMessage As String = "Failed to create new record."
    Public NumericFieldMessage As String = "Please enter numbers for this field."
    Public AlphaNumericFieldMessage As String = "Please enter numbers or alphabet letters for this field."
    Public Shared CurrencyFormat As String = "###,###,###,##0.00"
    Public DefaultDateTimeFormat As String = "dd-MMM-yyyy hh:mm tt"
    Public DefaultDateFormat As String = "dd-MMM-yyyy"
    Public PrintDateFormat As String = "dd MMMM yyyy"
    Public PasswordResetSuccessfully As String = "Password reset successfully."
    Public Shared ClassificationString = "P|Personal;;B|Business;;G|General;;P|Promotional;;S|Social;;F|Friends;;O|Others"


    Public Function AbsoluteWebPath() As String
        Dim domain As String = ""
        Dim domainPart
        domainPart = HttpContext.Current.Request.Url.ToString().Replace(HttpContext.Current.Request.Url.Scheme + "://", "").Split("/".ToCharArray())

        If (domainPart.Length > 0) Then
            domain = HttpContext.Current.Request.Url.Scheme + "://" + domainPart(0)
        End If

        Dim returnURL As String = "http://www.humecementconnect.com.my"
        '        returnURL = (domain + HttpContext.Current.Request.RawUrl).Replace(WebLib.GetFileNameWithQueryString(), "")

        Return returnURL
    End Function
    Public Shared Function GetRightsByProfileID(ByVal ProfileID As String) As String

        Dim connectionstring As String = System.Configuration.ConfigurationSettings.AppSettings("ConnStr")

        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim dr As DataRow
        Dim ltemp As String = ""

        ltemp = "SELECT CAST((select ltrim('''' + case rtrim(ssa_add) when '' then  '' else ssa_functioncode + ssa_add + ';;' end + case rtrim(ssa_mod) when '' then  '' else ssa_functioncode + ssa_mod + ';;' end  +  case rtrim(ssa_del) when '' then  '' else ssa_functioncode + ssa_del + ';;' end +  case rtrim(ssa_view) when '' then  '' else ssa_functioncode + ssa_view + ';;' end +case rtrim(ssa_full) when '' then  '' else ssa_functioncode + ssa_full + ';;' end  + ''',') as 'data()' from secaccessrights Where  ssa_filtercode='" & Weblib.Filtercode & "' and ssa_profileid=" & ProfileID & " for xml path('')) AS VARCHAR(MAX)) AS RtnData"

        cn.open()
        cmd.CommandText = ltemp
        cmd.Connection = cn
        ad.SelectCommand = cmd
        ad.Fill(ds, "datarecords")
        For Each dr In ds.Tables("datarecords").Rows
            counter = counter + 1
            LoginRightsString = dr("RtnData") & ""
        Next

        cn.Close()
        cmd.dispose()
        cn.dispose()

    End Function
    Public Shared Function ClientURL(ByVal pURL As String) As String
        Dim p As Page = CType(System.Web.HttpContext.Current.Handler, Page)
        Return p.ResolveClientUrl("~/" & pURL)
    End Function

    Public Shared Property PrevURL()
        Get
            If (HttpContext.Current.Session("PrevURL") & "" <> "") Then

                Return HttpContext.Current.Session("PrevURL")
            Else
                Return ""
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("PrevURL") = value

        End Set
    End Property
    Public Shared Function getclassificationImage(ByVal pClassification As String)
        Dim limagefile As String = "others.png"
        Select Case pClassification.trim
            Case "P"
                limagefile = "personal.png"
            Case "G"
                limagefile = "general.png"
            Case "B"
                limagefile = "business.png"
            Case "S"
                limagefile = "social.png"
            Case "F"
                limagefile = "friends.png"
            Case "M"
                limagefile = "promotional.png"

            Case Else
                limagefile = "others.png"
        End Select

        Return ClientURL("graphics/classification/" & limagefile)

    End Function

    Public Shared Function SendSMS(ByVal pMessage As String, ByVal pHPcc As String, ByVal pHPNo As String)

        Return True
        
    End Function
    Public Shared Property LoginUserHPCC()
        Get
            If (HttpContext.Current.Session("LoginUserHPCC") & "" <> "") Then

                Return HttpContext.Current.Session("LoginUserHPCC")
            Else
                Return ""
            End If

        End Get
        Set(ByVal value)
            HttpContext.Current.Session("LoginUserHPCC") = value
        End Set
    End Property
    Public Shared Function GetClickkMePath() As String
        Return "http://clickk.me/"
    End Function
    Public Shared Function InitClickkMe(ByVal _p_Type As String, ByVal _p_ID As String, Optional ByVal pDaystoExpiry As Integer = 3, Optional ByVal Noofclicks As Integer = 2) As String
        Return ""

    End Function


    Public Shared Property LoginUserHPNo()
        Get
            If (HttpContext.Current.Session("LoginUserHPNo") & "" <> "") Then

                Return HttpContext.Current.Session("LoginUserHPNo")
            Else
                Return ""
            End If

        End Get
        Set(ByVal value)
            HttpContext.Current.Session("LoginUserHPNo") = value
        End Set
    End Property
    Public Shared Function getUniqueKey()

        Dim obj As New ViFeandi.General
        Return obj.getUniqueCode(6)
        obj = Nothing
        '        Return (merchantid & filtercode & loginuser & DateTime.Now.ToString.Replace(" ", "")).ToLower
    End Function

    Public Shared Property LoginUserEmail()
        Get
            If (HttpContext.Current.Session("LoginUserEmail") & "" <> "") Then

                Return HttpContext.Current.Session("LoginUserEmail")
            Else
                Return ""
            End If

        End Get
        Set(ByVal value)
            HttpContext.Current.Session("LoginUserEmail") = value
        End Set
    End Property

    Public Shared Property LoginRightsString()
        Get
            If (HttpContext.Current.Session("LoginRightsString") & "" <> "") Then

                Return HttpContext.Current.Session("LoginRightsString")
            Else
                Return ""
            End If

        End Get
        Set(ByVal value)
            HttpContext.Current.Session("LoginRightsString") = value
        End Set
    End Property
    Public Shared Property PrintEngine()
        Get
            'Return "http://humecementconnect.com.my/printengine/"

            'Return "http://219.93.25.100/printengine/"

            Return "http://www.humecementconnect.com.my/printengine/"

        End Get
        Set(ByVal value)

        End Set
    End Property

    Public Shared Property ConnFMS()
        Get
			' 20201102
            'Return "Provider=SQLOLEDB;Data Source=132.18.10.220;Initial Catalog=FMS_Staging;User ID=visoft;Password=f6Zu7HcN98PLzqP"
            Return "Provider=SQLOLEDB;Data Source=132.18.10.106;Initial Catalog=FMS_Staging;User ID=visoft;Password=f6Zu7HcN98PLzqP"

        End Get
        Set(ByVal value)

        End Set
    End Property
    Public Shared Property ConnEpicor()
        Get
            Return "dsn=HCMT-SQL;uid=USR_PORTAL;pwd=1qazXSW@#;"
        End Get
        Set(ByVal value)

        End Set
    End Property
    Public Shared Property ConnAD1()
        Get
            Return "PJ-DC01.hume.net" 	'132.6.10.253
        End Get
        Set(ByVal value)

        End Set
    End Property
    Public Shared Property ConnAD2()
        Get
            Return "PJ-SVR01.hume.net" 	'132.6.10.250
        End Get
        Set(ByVal value)

        End Set
    End Property
    Public Shared Property ConnAD3()
        Get
            Return "PJ-DC02.hume.net"	'132.6.10.202
        End Get
        Set(ByVal value)

        End Set
    End Property
    Public Shared Property ConnAD4()
        Get
            Return "PJ-DC03.hume.net"	'132.6.10.237
        End Get
        Set(ByVal value)

        End Set
    End Property
    Public Shared Function GetTime(ByVal pHHMM As String, ByVal pHourOrMinute As String) As Integer
        Dim ltemp As String = ""
        Select Case pHourOrMinute.tolower

            Case "hour"
                ltemp = Microsoft.VisualBasic.Left(pHHMM, 2)
            Case "minute"

                ltemp = Microsoft.VisualBasic.Right(pHHMM, 2)
            Case Else
                ltemp = 0
        End Select

        If isnumeric(ltemp) = False Then
            Return 0
        Else
            Return CLng(ltemp)
        End If

    End Function
    Public Shared Property CustBranchID()
        Get
            If (HttpContext.Current.Session("CustBranchID") & "" <> "") Then

                Return HttpContext.Current.Session("CustBranchID")
            Else
                Return ""
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("CustBranchID") = value

        End Set
    End Property
    Public Shared Property CustBranchNum()
        Get
            If (HttpContext.Current.Session("CustBranchNum") & "" <> "") Then

                Return HttpContext.Current.Session("CustBranchNum")
            Else
                Return ""
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("CustBranchNum") = value

        End Set
    End Property
    Public Shared Property CustBranchName()
        Get
            If (HttpContext.Current.Session("CustBranchName") & "" <> "") Then

                Return HttpContext.Current.Session("CustBranchName")
            Else
                Return ""
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("CustBranchName") = value

        End Set
    End Property
    Public Shared Property MerchantID()
        Get
            If (HttpContext.Current.Session("MerchantID") & "" <> "") Then

                Return HttpContext.Current.Session("MerchantID")
            Else
                Return ""
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("MerchantID") = value

        End Set
    End Property
    Public Shared Function ExportExcelODBC(ByVal Response As Object, ByVal pSQLStatement As String, ByVal pFilePrefix As String) As Boolean


        Dim cn As New OdbcConnection(Weblib.connEpicor)
        Dim cmd As OdbcCommand
        Dim ad As New Odbc.OdbcDataAdapter


        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim tw As New System.IO.StringWriter()
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)
        Dim dgGrid As DataGrid = New DataGrid
        If pFileprefix.trim = "" Then
            pFileprefix = "Export"
        End If
        Dim filename As String = pFilePrefix & Microsoft.VisualBasic.Format(datetime.today.year, "00") & Microsoft.VisualBasic.Format(datetime.today.month, "00") & Microsoft.VisualBasic.Format(datetime.today.day, "00") & ".xls"
        Try

            cn.Open()


            cmd = New OdbcCommand(pSQLStatement, cn)
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")

            Dim dt As DataTable = ds.Tables("datarecords")


            dggrid.HeaderStyle.ForeColor = Drawing.Color.White
            dggrid.HeaderStyle.BackColor = Drawing.Color.Black
            dggrid.ItemStyle.VerticalAlign = VerticalAlign.Top
            dggrid.ItemStyle.HorizontalAlign = HorizontalAlign.Left

            dggrid.ItemStyle.cssclass = "textmode"

            dgGrid.DataSource = dt
            dgGrid.DataBind()


            dgGrid.RenderControl(hw)

            Response.Clear()
            Response.Buffer = True
            Response.AddHeader("content-disposition", "attachment;filename=" & filename & "")
            Response.Charset = ""
            Response.ContentType = "application/vnd.ms-excel"
            Response.Output.Write(tw.ToString())
            Response.Flush()
            Response.End()
            Return True
        Catch ex As Exception
            Weblib.ErrorTrap = ex.message
            Return False
        End Try

    End Function
    Public Shared Function ExportExcel(ByVal Response As Object, ByVal pSQLStatement As String, ByVal pFilePrefix As String, Optional ByVal connstring As String = "") As Boolean
        Dim lconn As String = ""
        If connstring.trim <> "" Then
            lconn = connstring
        Else
            lconn = System.Configuration.ConfigurationSettings.AppSettings("ConnStr")
        End If

        Dim cn As New OleDbConnection(lconn)

        Dim cmd As New OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim tw As New System.IO.StringWriter()
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)
        Dim dgGrid As DataGrid = New DataGrid
        If pFileprefix.trim = "" Then
            pFileprefix = "Export"
        End If
        Dim filename As String = pFilePrefix & Microsoft.VisualBasic.Format(datetime.today.year, "00") & Microsoft.VisualBasic.Format(datetime.today.month, "00") & Microsoft.VisualBasic.Format(datetime.today.day, "00") & ".xls"
        Try

            cn.Open()
            cmd.CommandText = pSQLStatement
            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")
            Dim dt As DataTable = ds.Tables("datarecords")


            dggrid.HeaderStyle.ForeColor = Drawing.Color.White
            dggrid.HeaderStyle.BackColor = Drawing.Color.Black
            dggrid.ItemStyle.VerticalAlign = VerticalAlign.Top
            dggrid.ItemStyle.HorizontalAlign = HorizontalAlign.Left

            dggrid.ItemStyle.cssclass = "textmode"

            dgGrid.DataSource = dt
            dgGrid.DataBind()


            dgGrid.RenderControl(hw)

            Response.Clear()
            Response.Buffer = True
            Response.AddHeader("content-disposition", "attachment;filename=" & filename & "")
            Response.Charset = ""
            Response.ContentType = "application/vnd.ms-excel"
            Response.Output.Write(tw.ToString())
            Response.Flush()
            Response.End()
            Return True
        Catch ex As Exception
            Weblib.ErrorTrap = ex.message
            Return False
        End Try

    End Function

    Public Shared Property CustNum()
        Get
            If (HttpContext.Current.Session("CustNum") & "" <> "") Then

                Return HttpContext.Current.Session("CustNum")
            Else
                Return ""
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("CustNum") = value

        End Set
    End Property

    Public Shared Property CustName()
        Get
            If (HttpContext.Current.Session("CustName") & "" <> "") Then

                Return HttpContext.Current.Session("CustName")
            Else
                Return ""
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("CustName") = value

        End Set
    End Property

    Public Shared Property CustCode()
        Get
            If (HttpContext.Current.Session("CustCode") & "" <> "") Then

                Return HttpContext.Current.Session("CustCode")
            Else
                Return ""
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("CustCode") = value

        End Set
    End Property
    Public Shared Property ApplicationID()
        Get
            Return "AppID"
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("AppID") = value
        End Set
    End Property
    Public Shared Property ErrorTrap()
        Get
            If (HttpContext.Current.Session("ErrorTrap") & "" <> "") Then

                Return HttpContext.Current.Session("ErrorTrap")
            Else
                Return ""
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("ErrorTrap") = value

        End Set
    End Property

    Public Shared Property FilterCode()
        Get
            Return "Filter"
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("FilterCode") = value
        End Set
    End Property
    Public Shared Property ProfileID()
        Get
            If (HttpContext.Current.Session("ProfileID") & "" <> "") Then

                Return HttpContext.Current.Session("ProfileID")
            Else
                Return ""
            End If

        End Get
        Set(ByVal value)
            HttpContext.Current.Session("ProfileID") = value
        End Set
    End Property
    Public Shared Property BranchID()
        Get
            If (HttpContext.Current.Session("BranchID") & "" <> "") Then

                Return HttpContext.Current.Session("BranchID")
            Else
                Return ""
            End If

        End Get
        Set(ByVal value)
            HttpContext.Current.Session("BranchID") = value
        End Set
    End Property
    Public Shared Property AppCodeString()
        Get
            If (HttpContext.Current.Session("AppCodeString") & "" <> "") Then

                Return HttpContext.Current.Session("AppCodeString")
            Else
                Return ""
            End If

        End Get
        Set(ByVal value)
            HttpContext.Current.Session("AppCodeString") = value
        End Set
    End Property
    Public Shared Property LoginUser()
        Get
            If (HttpContext.Current.Session("LoggedInUser") & "" <> "") Then

                Return HttpContext.Current.Session("LoggedInUser")
            Else
                Return ""
            End If

        End Get
        Set(ByVal value)
            HttpContext.Current.Session("LoggedInUser") = value
        End Set
    End Property
    Public Shared Function getAlertMessageStyle(ByVal pMessage As String) As String
        Return "<div class=""ui-state-highlight"" style=""padding-left:10px; padding-top:10px; padding-right:10px; padding-bottom:10px;width:100%""><span class=""ui-icon ui-icon-alert"" style=""float: left; margin-right:.3em;""></span>" & pMessage & "</div>"
    End Function

    Public Shared Function ShowLoginInfo() As String
        If (Weblib.merchantid).trim = "" Then
            Return "<div class=""ui-state-highlight"" style=""padding-left:10px; padding-top:10px; padding-right:10px; padding-bottom:10px;width:100%""><font class=""cssdetail"">Login Company : <font color=""red"">NOT SELECTED</font></font><font class=""cssdetail"">&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;Login User : " & weblib.loginusername & "</font><font class=""cssdetail"">&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;Branch : Not Applicable</font></div>"

        Else
            If (weblib.CustNum & "").trim = "" Then
                Dim obj As New RuntimeCustomer
                Call obj.getinfo(Weblib.MerchantID)
                weblib.CustNum = obj.CustNum
                weblib.CustName = obj.CustName
                obj = Nothing
            End If

            If (weblib.custbranchid & "").trim = "" Then
                weblib.custbranchname = "Not Applicable"
            End If
            Return "<div class=""ui-state-highlight"" style=""padding-left:10px; padding-top:10px; padding-right:10px; padding-bottom:10px;width:100%""><font class=""cssdetail"">Login Company : </font><font class=""cssdetail""><b>" & weblib.CustName & "</b>&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;Login User : " & weblib.loginusername & "</font><font class=""cssdetail"">&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;Branch : " & weblib.custbranchname & "</font></div>"
        End If
    End Function


    Public Shared Property isStaff() As Boolean
        Get

            Return weblib.bittoboolean(HttpContext.Current.Session("isStaff"))

        End Get
        Set(ByVal value As Boolean)
            HttpContext.Current.Session("isStaff") = value
        End Set
    End Property

    Public Shared Property LoginUserName()
        Get
            If (HttpContext.Current.Session("LoggedInUserName") & "" <> "") Then

                Return HttpContext.Current.Session("LoggedInUserName")
            Else
                Return ""
            End If

        End Get
        Set(ByVal value)
            HttpContext.Current.Session("LoggedInUserName") = value
        End Set
    End Property
    Public Shared Property StartupApp()
        Get
            If (HttpContext.Current.Session("StartupAPP") & "" <> "") Then

                Return HttpContext.Current.Session("StartupAPP")
            Else
                Return ""
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("StartupAPP") = value

        End Set
    End Property
    Public Shared Function DefaultHumanPng() As String
        Return "images/default.png"
    End Function

    Public Shared Function BitToBoolean(ByVal _p_value As String) As Boolean

        If IsNumeric(_p_value) = False Then
            Select Case LCase(_p_value)
                Case "true"
                    Return True
                Case "false"
                    Return False
                Case Else
                    Return False
            End Select
        Else
            If CInt(_p_value) = 0 Then
                Return False
            Else
                Return True
            End If
        End If

    End Function
    Public Shared Function BooleanToBit(ByVal _p_value As Boolean) As Integer
        If _p_value = True Then
            Return 1
        Else
            Return 0
        End If
    End Function
    Public Shared Sub ShowMessagePage(ByRef pResponse As Object, ByVal pMessage As String, ByVal pNextPage As String)

        pResponse.Redirect(ClientURL("postpage.aspx?NextPage=message.aspx&ga=" & pMessage & "&ba=" & pNextPage))

    End Sub
    Public Shared Function RandomString(ByVal pSize As Integer, ByVal lowercase As Boolean) As String
        Dim builder As New stringbuilder()
        Dim random As New Random()

        Dim ch As Char

        Dim i As Integer
        For i = 0 To psize
            ch = Convert.ToChar(Convert.ToInt32(Math.Floor(26 * random.NextDouble() + 65)))
            builder.append(ch)
        Next
        If lowercase = True Then
            Return builder.ToString().ToLower()
        Else
            Return builder.ToString()
        End If

    End Function
    Public Shared Function CodeExists(ByVal pCode As String, ByVal pFieldName As String, ByVal pTable As String, ByVal pIDField As String, ByVal pID As String, ByVal pIDPField As String, ByVal pIDP As String, ByVal pAppID As String, ByVal pMerchantID As String, ByVal pFilterCode As String, Optional ByVal pSQL As String = "") As Boolean
        Dim connectionstring As String = System.Configuration.ConfigurationSettings.AppSettings("ConnStr")

        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim dr As DataRow
        Dim ltemp As String = ""
        If isnumeric(pID) = False Then

        Else
            ltemp = ltemp & " and " & pIDField & "<>" & pID
        End If

        If pIDPField.trim <> "" Then
            ltemp = ltemp & " and " & pIDPField & "=" & pIDP
        End If
        If pAppID.trim <> "" Then
            ltemp = ltemp & " and " & pAppID & "='" & WebLib.ApplicationID & "'"
        End If
        If pMerchantID.trim <> "" Then
            ltemp = ltemp & " and " & pMerchantID & "='" & WebLib.MerchantID & "'"
        End If
        If pFilterCode.trim <> "" Then
            ltemp = ltemp & " and " & pFilterCode & "='" & WebLib.FilterCode & "'"
        End If
        If psql.trim <> "" Then
            ltemp = ltemp & " and " & psql
        End If

        cn.open()
        cmd.CommandText = "Select top 1 " & pIDField & " from " & pTable & " where " & pFieldName & "='" & pCode & "' " & ltemp
        '        response.write(cmd.CommandText)
        weblib.ErrorTrap = cmd.CommandText
        cmd.Connection = cn
        ad.SelectCommand = cmd
        ad.Fill(ds, "datarecords")
        For Each dr In ds.Tables("datarecords").Rows
            counter = counter + 1
            Exit For
        Next

        If counter = 0 Then
            Return False
        Else
            Return True
        End If
        cn.Close()
    End Function
    Public Shared Function setListItemsTable(ByVal pObject As Object, ByVal pDisplayField As String, ByVal pValueField As String, ByVal pTable As String, ByVal pOrderbyField As String, ByVal pAppID As String, ByVal pMerchantID As String, ByVal pFilterCode As String, Optional ByVal pWhereField As String = "", Optional ByVal ShowPleaseSelect As Boolean = False) As Boolean
        Dim connectionstring As String = System.Configuration.ConfigurationSettings.AppSettings("ConnStrAdmin")

        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim dr As DataRow
        Dim ltemp As String = ""
        Dim lorderby As String = ""
        Dim item As ListItem

        Try

            If pORderbyField.trim <> "" Then
                lorderby = " Order by " & pORderbyField
            End If
            If pAppID.trim <> "" Then
                If ltemp.trim <> "" Then
                    ltemp = ltemp & " and "
                Else
                    ltemp = ltemp & " where "
                End If
                ltemp = ltemp & pAppID & "='" & WebLib.ApplicationID & "'"
            End If
            If pMerchantID.trim <> "" Then
                If ltemp.trim <> "" Then
                    ltemp = ltemp & " and "
                Else
                    ltemp = ltemp & " where "
                End If

                ltemp = ltemp & " " & pMerchantID & "='" & WebLib.MerchantID & "'"
            End If
            If pFilterCode.trim <> "" Then
                If ltemp.trim <> "" Then
                    ltemp = ltemp & " and "
                Else
                    ltemp = ltemp & " where "
                End If

                ltemp = ltemp & " " & pFilterCode & "='" & WebLib.FilterCode & "'"
            End If

            If pWhereField.trim <> "" Then
                If ltemp.trim <> "" Then
                    ltemp = ltemp & " and "
                Else
                    ltemp = ltemp & " where "
                End If

                ltemp = ltemp & " " & pWhereField


            End If

            If ShowPleaseSelect = True Then
                item = New ListItem("-Please Select-", "")
                pObject.Items.Add(item)

            End If

            cmd.CommandText = "Select " & pValueField & " as ValueField," & pDisplayField & " as DisplayField from " & pTable & " " & ltemp & lorderby
            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")
            For Each dr In ds.Tables("datarecords").Rows
                counter = counter + 1
                item = New ListItem(dr("DisplayField") & "", dr("ValueField") & "")
                pObject.Items.Add(item)
            Next
            cn.Close()
            cmd.dispose()
            cn.dispose()
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function
    Public Shared Function setListItemsTableEpicor(ByVal pObject As Object, ByVal pDisplayField As String, ByVal pValueField As String, ByVal pTable As String, ByVal pOrderbyField As String, Optional ByVal pWhereField As String = "", Optional ByVal pFirstItemBlank As Boolean = False) As Boolean


        Dim cn As New OdbcConnection(Weblib.connEpicor)
        Dim cmd As New OdbcCommand
        Dim ad As New Odbc.OdbcDataAdapter


        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim dr As DataRow
        Dim ltemp As String = ""
        Dim lorderby As String = ""
        Dim item As ListItem

        Try

            If pORderbyField.trim <> "" Then
                lorderby = " Order by " & pORderbyField
            End If

            If pWhereField.trim <> "" Then
                If ltemp.trim <> "" Then
                    ltemp = ltemp & " and "
                Else
                    ltemp = ltemp & " where "
                End If

                ltemp = ltemp & " " & pWhereField

            End If


            cmd.CommandText = "Select " & pValueField & " as ValueField," & pDisplayField & " as DisplayField from " & pTable & " " & ltemp & lorderby
            Weblib.errortrap = cmd.CommandText

            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")
            For Each dr In ds.Tables("datarecords").Rows
                counter = counter + 1
                item = New ListItem(dr("DisplayField") & "", dr("ValueField") & "")
                pObject.Items.Add(item)
            Next
            cn.Close()
            cmd.dispose()
            cn.dispose()


            If pFirstItemBlank = True Then
                pObject.Items.Insert(0, New ListItem("- Please Select -", ""))
                pObject.SelectedIndex = 0
            End If

            Return True
        Catch ex As Exception
            Weblib.errortrap = ex.message
            Return False
        End Try

    End Function

    Public Shared Function GetposDocNo(ByVal pType As String) As String

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

            ltemp = ltemp & " where doc_merchantid='" & WebLib.MerchantID & "'"
            ltemp = ltemp & " and doc_filter='" & WebLib.FilterCode & "'"
            ltemp = ltemp & " and doc_type='" & pType & "'"
            ltemp = ltemp & " and doc_branchid=" & WebLib.Branchid & ""

            cn.open()
            cmd.CommandText = "Select * from posDocumentNo " & ltemp
            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")
            For Each dr In ds.Tables("datarecords").Rows
                counter = counter + 1

                lDocNo = InitialValues(dr("doc_prefix") & "") & Microsoft.visualbasic.Format(dr("doc_no"), dr("doc_format")) & InitialValues(dr("doc_suffix") & "")

            Next

            If ldocno.trim = "" Then
                Return ""
                Exit Function
            End If

            cmd.CommandText = "Update posDocumentNo set doc_no = isnull(doc_no,0) + 1 " & ltemp
            cmd.ExecuteNonQuery()
            cn.Close()
            cmd.dispose()
            cn.dispose()
            Return ldocno
        Catch ex As Exception
            Return ""
        End Try

    End Function
    Public Shared Function formatthedate(ByVal pDate As Object, Optional ByVal MDY As Boolean = False) As String
        Try

            If microsoft.visualbasic.isdate(pDate) = True Then
                If mdy = True Then
                    Return CDate(pDate).ToString("MM/dd/yyyy")

                Else
                    Return CDate(pDate).ToString("dd/MM/yyyy")

                End If
                '                Return CDate(pDate).ToString("MM/dd/yyyy")

            Else
                Return ""
            End If


        Catch ex As Exception
            Return ex.message 'pdate.tostring
        End Try

    End Function
    Public Shared Function formatthemoney(ByVal pamount As String) As String
        Try

            If isnumeric(pAmount) = True Then

                Return Microsoft.VisualBasic.Format(CDbl(pamount), CurrencyFormat)

            Else
                Return Microsoft.VisualBasic.Format(CDbl("0"), CurrencyFormat)

            End If

        Catch ex As Exception
            Return Microsoft.VisualBasic.Format(CDbl("0"), CurrencyFormat)
        End Try

    End Function
    Public Shared Function formatthenumber(ByVal pamount As String, ByVal pFormat As String) As String
        Try

            If isnumeric(pAmount) = True Then

                Return Microsoft.VisualBasic.Format(CDbl(pamount), pFormat)

            Else
                Return Microsoft.VisualBasic.Format(CDbl("0"), pFormat)

            End If

        Catch ex As Exception
            Return Microsoft.VisualBasic.Format(CDbl("0"), pFormat)
        End Try

    End Function
    Public Shared Function GetDocNo(ByVal pNameSpace As String, ByVal pFieldName As String, ByVal pAppCode As String) As String

        Dim connectionstring As String = System.Configuration.ConfigurationSettings.AppSettings("ConnStr")

        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim dr As DataRow
        Dim ltemp As String = ""
        Dim lDocNo As String = ""
        '        Try

        'ltemp = ltemp & "where doc_appid='" & WebLib.ApplicationID & "'"
        ltemp = ltemp & "where doc_appid='" & pAppCode & "'"

        '        ltemp = ltemp & " and doc_merchantid='" & WebLib.MerchantID & "'"
        '       ltemp = ltemp & " and doc_filter='" & WebLib.FilterCode & "'"
        ltemp = ltemp & " and doc_namespace='" & pNameSpace & "'"
        ltemp = ltemp & " and doc_fieldname='" & pFieldName & "'"

        cn.open()
        cmd.CommandText = "Select * from DocumentNo with (nolock) " & ltemp
        cmd.Connection = cn
        ad.SelectCommand = cmd
        ad.Fill(ds, "datarecords")
        For Each dr In ds.Tables("datarecords").Rows
            counter = counter + 1

            lDocNo = InitialValues(dr("doc_prefix") & "") & Microsoft.visualbasic.Format(dr("doc_no"), dr("doc_format")) & InitialValues(dr("doc_suffix") & "")

        Next

        cmd.CommandText = "Update DocumentNo set doc_no = isnull(doc_no,0) + 1 " & ltemp
        cmd.ExecuteNonQuery()

        cn.Close()
        cmd.dispose()
        cn.dispose()
        Return lDocno
        '        Catch ex As Exception
        'Return ""
        'End Try

    End Function
    Public Shared Property LoginIsFullAdmin() As Boolean
        Get
            If (HttpContext.Current.Session("LoginIsFullAdmin") & "" <> "") Then

                Return HttpContext.Current.Session("LoginIsFullAdmin")
            Else
                Return False
            End If

        End Get
        Set(ByVal value As Boolean)
            HttpContext.Current.Session("LoginIsFullAdmin") = value
        End Set
    End Property
    Public Shared Property FilterPartClass() As String
        Get
            'Return "'C211','C222','C212','C223','C213','C221','C231','GGBFS','PCC BULK'"
            Return "'C211','C222','C212','C223','C213','C221','C231','GGBFS','PCC BULK','C242','C243','C252','C253','C251'"

        End Get
        Set(ByVal value As String)

        End Set
    End Property
    Private Shared Function InitialValues(ByVal pString As String) As String
        pString = pString.replace("YYYY", Microsoft.VisualBasic.Format(today.year, "0000"))
        pString = pString.replace("YY", Microsoft.VisualBasic.Format(today.year, "00"))
        pString = pString.replace("MM", Microsoft.VisualBasic.Format(today.month, "00"))
        pString = pString.replace("DD", Microsoft.VisualBasic.Format(today.day, "00"))
        Return pString
    End Function
    Public Shared Sub SetListItems(ByVal pObject As Object, ByVal pData As String)
        Dim temp
        Dim temp2

        Dim item As ListItem
        Dim counter As Integer
        Dim lstring = pData
        If lstring <> "" Then
            temp = Microsoft.VisualBasic.Split(lstring, ";;")

            If Microsoft.VisualBasic.UBound(temp) >= 1 Then
                For counter = 0 To Microsoft.VisualBasic.UBound(temp)
                    temp2 = Microsoft.VisualBasic.Split(temp(counter), "|")
                    If Microsoft.VisualBasic.UBound(temp2) >= 1 Then
                        item = New ListItem(temp2(1), temp2(0))
                        pObject.Items.Add(item)
                    Else
                        item = New ListItem(temp(counter), temp(counter))
                        pObject.Items.Add(item)
                    End If
                Next
            Else

                temp2 = Microsoft.VisualBasic.Split(lstring, "|")
                If Microsoft.VisualBasic.UBound(temp2) >= 1 Then
                    item = New ListItem(temp2(0), temp2(1))
                    pObject.Items.Add(item)
                Else
                    item = New ListItem(lstring, lstring)
                    pObject.Items.Add(item)
                End If

            End If

        End If

    End Sub
    Public Shared Function hasrightsApp(ByVal pAppCode As String) As Boolean

        Return True

    End Function
    Public Shared Function GetAppsByMerchantIDforSearch(ByVal pMerchantID As String) As String

        Dim connectionstring As String = System.Configuration.ConfigurationSettings.AppSettings("ConnStr")

        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim dr As DataRow
        Dim ltemp As String = ""

        ltemp = "SELECT CAST((select ltrim('''' + mr_AppCode + ''',') as 'data()' from sysMerchantAPP where mr_merchantid='" & pMerchantID & "' for xml path('')) AS VARCHAR(MAX)) AS RtnData"

        cn.open()
        cmd.CommandText = ltemp
        cmd.Connection = cn
        ad.SelectCommand = cmd
        ad.Fill(ds, "datarecords")
        For Each dr In ds.Tables("datarecords").Rows
            counter = counter + 1
            ltemp = dr("RtnData") & ""
        Next
        If counter = 0 Then
            ltemp = ""
        End If

        cn.Close()
        cmd.dispose()
        cn.dispose()

        Return ltemp
    End Function
    Public Shared Sub GetAppsByMerchantID(ByVal pMerchantID As String)

        Dim connectionstring As String = System.Configuration.ConfigurationSettings.AppSettings("ConnStr")

        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim dr As DataRow
        Dim ltemp As String = ""

        ltemp = "SELECT CAST((select ltrim(mr_AppCode + ';;') as 'data()' from sysMerchantAPP where mr_merchantid='" & pMerchantID & "' for xml path('')) AS VARCHAR(MAX)) AS RtnData"

        cn.open()
        cmd.CommandText = ltemp
        cmd.Connection = cn
        ad.SelectCommand = cmd
        ad.Fill(ds, "datarecords")
        For Each dr In ds.Tables("datarecords").Rows
            counter = counter + 1
            AppCodeString = dr("RtnData") & ""
        Next

        cn.Close()
        cmd.dispose()
        cn.dispose()

    End Sub
    Public Shared Function hasmodrights(ByVal pNameSpace As String, ByVal pAppCode As String, ByVal pPrefix As String) As Boolean

        If hasrights(pNAmeSpace, pAppCode, pPrefix & "1") = True Then
            Return True
            Exit Function
        End If

        If hasrights(pNAmeSpace, pAppCode, pPrefix & "2") = True Then
            Return True
            Exit Function
        End If

        If hasrights(pNAmeSpace, pAppCode, pPrefix & "3") = True Then
            Return True
            Exit Function
        End If

        If hasrights(pNAmeSpace, pAppCode, pPrefix & "4") = True Then
            Return True
            Exit Function
        End If

        If hasrights(pNAmeSpace, pAppCode, pPrefix & "5") = True Then
            Return True
            Exit Function
        End If
        Return False

    End Function
    Public Shared Function hasrights(ByVal pNameSpace As String, ByVal pAppCode As String, ByVal pRights As String) As Boolean
        If loginisfulladmin = True Then
            Return True
            Exit Function
        End If
        If (pNameSpace & "" & pRights & "").trim = "" Then
            Return False
            Exit Function
        End If
        If microsoft.visualbasic.instr(1, LoginRightsString, pNameSpace & pRights & ";;") > 0 Then
            Return True
        Else
            Return False
        End If


        Return True
    End Function
    Public Shared Sub SetItemsDataType(ByVal pObject As DropDownList)

        Dim item As ListItem
        item = New ListItem("Varchar", "varchar")
        pObject.Items.Add(item)
        item = New ListItem("Integer", "int")
        pObject.Items.Add(item)
        item = New ListItem("Decimal", "decimal")
        pObject.Items.Add(item)
        item = New ListItem("Bit", "bit")
        pObject.Items.Add(item)
        item = New ListItem("DateTime", "datetime")
        pObject.Items.Add(item)
        item = New ListItem("Text", "text")
        pObject.Items.Add(item)
        item = New ListItem("Identity", "identity")
        pObject.Items.Add(item)
        item = New ListItem("Parent Identity", "pidentity")
        pObject.Items.Add(item)
        item = New ListItem("Link Field", "linkfield")
        pObject.Items.Add(item)
    End Sub
    Public Shared Sub SetItemsObjectType(ByVal pObject As DropDownList)

        Dim item As ListItem

        item = New ListItem("Textbox", "textbox")
        pObject.Items.Add(item)
        item = New ListItem("Textbox Email", "textboxe")
        pObject.Items.Add(item)
        item = New ListItem("Radio Button List", "radio")
        pObject.Items.Add(item)
        item = New ListItem("Check Box", "checkbox")
        pObject.Items.Add(item)
        item = New ListItem("Dropdown", "dropdown")
        pObject.Items.Add(item)
        item = New ListItem("Password", "password")
        pObject.Items.Add(item)
        item = New ListItem("Multiline Text", "multiline")
        pObject.Items.Add(item)
        item = New ListItem("Date Picker", "datepicker")
        pObject.Items.Add(item)
        item = New ListItem("Upload Image", "fileimage")
        pObject.Items.Add(item)
        item = New ListItem("Upload Docs", "filefile")
        pObject.Items.Add(item)
        item = New ListItem("No Entry", "No")
        pObject.Items.Add(item)
        item = New ListItem("Create Date", "cgetdate")
        pObject.Items.Add(item)
        item = New ListItem("Create By", "cuserid")
        pObject.Items.Add(item)
        item = New ListItem("Update Date", "ugetdate")
        pObject.Items.Add(item)
        item = New ListItem("Update By", "uuserid")
        pObject.Items.Add(item)
        item = New ListItem("Filter Code", "filter")
        pObject.Items.Add(item)
        item = New ListItem("Application ID", "appid")
        pObject.Items.Add(item)
        item = New ListItem("Merchant ID", "merchantid")
        pObject.Items.Add(item)
        item = New ListItem("Branch ID", "branchid")
        pObject.Items.Add(item)
        item = New ListItem("Auto Code", "autocode")
        pObject.Items.Add(item)

    End Sub
    Public Shared Function GetValue(ByVal pTableName As String, ByVal pFieldName As String, ByVal pSearchFieldName As String, ByVal pSearchValue As String, ByVal pMerchantIDField As String, ByVal pFilterField As String, Optional ByVal pWhereField As String = "") As String

        Dim connectionstring As String = System.Configuration.ConfigurationSettings.AppSettings("ConnStr")

        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim dr As DataRow
        Dim ltemp As String = ""
        Dim lRtnValue As String = ""
        Try

            ltemp = ltemp & "where " & pSearchFieldName & "=" & pSearchValue & ""
            If pMerchantIDField.trim <> "" Then
                ltemp = ltemp & " and " & pMerchantIDField & "='" & WebLib.MerchantID & "'"
            End If
            If pFilterField.trim <> "" Then
                ltemp = ltemp & " and " & pFilterField & "='" & WebLib.FilterCode & "'"
            End If
            If pWhereField.trim <> "" Then
                If ltemp.trim <> "" Then
                    ltemp = ltemp & " and "
                Else
                    ltemp = ltemp & " where "
                End If
                ltemp = ltemp & " " & pWhereField
            End If

            cn.open()
            cmd.CommandText = "Select " & pFieldName & " as rtnvalue from " & pTableName & " " & ltemp
            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")
            For Each dr In ds.Tables("datarecords").Rows
                counter = counter + 1
                lRtnValue = dr("rtnvalue") & ""
            Next
            cn.Close()
            cmd.dispose()
            cn.dispose()
            Return lRtnValue
        Catch ex As Exception
            Return ""
        End Try

    End Function
    Public Shared Function GetValueClear(ByVal pTableName As String, ByVal pFieldName As String, ByVal pSearchFieldName As String, ByVal pSearchValue As String) As String

        Dim connectionstring As String = System.Configuration.ConfigurationSettings.AppSettings("ConnStr")

        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim dr As DataRow
        Dim ltemp As String = ""
        Dim lRtnValue As String = ""
        Try

            ltemp = ltemp & "where " & pSearchFieldName & "=" & pSearchValue & ""

            cn.open()
            cmd.CommandText = "Select " & pFieldName & " as rtnvalue from " & pTableName & " " & ltemp
            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")
            For Each dr In ds.Tables("datarecords").Rows
                counter = counter + 1
                lRtnValue = dr("rtnvalue") & ""
            Next
            cn.Close()
            cmd.dispose()
            cn.dispose()
            Return lRtnValue
        Catch ex As Exception
            Return ""
        End Try

    End Function
    Public Shared Property LoginUserCompanySelected()
        Get
            If (HttpContext.Current.Session("LoginUserCompanySelected") & "" <> "") Then

                Return HttpContext.Current.Session("LoginUserCompanySelected")
            Else
                Return ""
            End If

        End Get
        Set(ByVal value)
            HttpContext.Current.Session("LoginUserCompanySelected") = value
        End Set
    End Property
    Public Shared Property LoginUserMatrixLevel()
        Get
            If (HttpContext.Current.Session("LoginUserMatrixLevel") & "" <> "") Then

                Return HttpContext.Current.Session("LoginUserMatrixLevel")
            Else
                Return ""
            End If

        End Get
        Set(ByVal value)
            HttpContext.Current.Session("LoginUserMatrixLevel") = value
        End Set
    End Property
    Public Shared Property LoginUserRegion()
        Get
            If (HttpContext.Current.Session("LoginUserRegion") & "" <> "") Then

                Return HttpContext.Current.Session("LoginUserRegion")
            Else
                Return ""
            End If

        End Get
        Set(ByVal value)
            HttpContext.Current.Session("LoginUserRegion") = value
        End Set
    End Property
    Public Shared Property LoginUserState()
        Get
            If (HttpContext.Current.Session("LoginUserState") & "" <> "") Then

                Return HttpContext.Current.Session("LoginUserState")
            Else
                Return ""
            End If

        End Get
        Set(ByVal value)
            HttpContext.Current.Session("LoginUserState") = value
        End Set
    End Property
    Public Shared Property isAD()
        Get
            If (HttpContext.Current.Session("isAD") & "" <> "") Then

                Return HttpContext.Current.Session("isAD")
            Else
                Return ""
            End If

        End Get
        Set(ByVal value)
            HttpContext.Current.Session("isAD") = value
        End Set
    End Property
    Public Shared Property CustUnderLoginUserMatrixLevel()
        Get
            If (HttpContext.Current.Session("CustUnderLoginUserMatrixLevel") & "" <> "") Then

                Return HttpContext.Current.Session("CustUnderLoginUserMatrixLevel")
            Else
                Return ""
            End If

        End Get
        Set(ByVal value)
            HttpContext.Current.Session("CustUnderLoginUserMatrixLevel") = value
        End Set
    End Property

End Class
Public Class Record_SearckKey
    Public KeyCaption As String = ""
    Public KeyFieldName As String = ""
    Public KeyDataType As String = ""

    Function BreakKey(ByVal _p_SearchString As String) As Boolean
        Dim lstr
        lstr = Microsoft.VisualBasic.Split(_p_SearchString, ";")
        If Microsoft.VisualBasic.UBound(lstr) = 2 Then
            KeyCaption = lstr(1)
            KeyFieldName = lstr(0)
            KeyDataType = lstr(2)
            Return True
        Else
            Return False

        End If
    End Function
    Function DefineSearchKey(ByVal _p_String As String, ByVal _p_Value As String, ByVal _p_Value2 As String, Optional ByVal _p_connection As String = "") As String
        Dim lstr
        lstr = Microsoft.VisualBasic.Split(_p_String, ";")
        Dim str2 As String = ""
        If Microsoft.VisualBasic.UBound(lstr) = 2 Then
            KeyCaption = lstr(1)
            KeyFieldName = lstr(0)
            KeyDataType = lstr(2)

            Select Case KeyDataType
                Case "S"
                    Return KeyFieldName & " like '%" & _p_Value & "%'"
                Case "N"
                    If isnumeric(_p_Value) = True Then
                        Return KeyFieldName & "=" & _p_Value
                    Else
                        Return KeyFieldName & "=0"
                    End If
                Case "D"
                    '                    If isdate(_p_Value) = True And isdate(_p_value2) = True Then


                    If _p_connection.tolower = "epicor" Then
                        If _p_Value2 <> "" Then
                            str2 = " and " & KeyFieldName & " <= '" & _p_Value2 & "' "
                        End If
                        'Return " (" & KeyFieldName & " >= '" & _p_Value & "' and " & KeyFieldName & " <= '" & _p_Value2 & "') "
                        Return " (" & KeyFieldName & " >= '" & _p_Value & "' " & str2 & ") "

                    Else
                        If _p_Value2 <> "" Then
                            str2 = " and datediff(d,'" & _p_Value2 & "'," & KeyFieldName & ") <=0 "
                        End If
                        'Return " (datediff(d,'" & _p_Value & "'," & KeyFieldName & ") > = 0 and datediff(d,'" & _p_Value2 & "'," & KeyFieldName & ") <=0 )"
                        Return " (datediff(d,'" & _p_Value & "'," & KeyFieldName & ") > = 0 " & str2 & " )"

                    End If



                    'Else
                    'Return ""
                    'End If
                Case Else
                    Return ""
            End Select
        Else

            Return ""
        End If
    End Function

End Class






