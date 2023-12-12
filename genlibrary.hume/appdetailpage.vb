Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System
Imports System.Configuration
Imports System.Collections.Generic
Imports System.Xml

Public Class appdetailspage

    Inherits System.Web.UI.Page
    Public connectionstring As String = System.Configuration.ConfigurationSettings.AppSettings("ConnStr")
    Public listingpage As String = ""
    Public _FormsName As String = ""
    Public TableName As String = ""
    Public DetailPage As String = ""
    Public IDPField As String = ""
    Public IDField As String = ""
    Public Orderby As String = ""
    Public AppIDField As String = ""
    Public MerchantIDField As String = ""
    Public FilterField As String = ""
    Public bid, rid, lblmessage, btnsave, btnback As Object
    Public _Mode As String = ""
    Public fUserID As String = ""
    Public APPCode As String = ""
    Public AddRights As String = ""
    Public DelRights As String = ""
    Public ModRights As String = ""
    Public ViewRights As String = ""
    Public FullRights As String = ""
    Public NmSpace As String = ""
    Public isSkipMerchatID As Boolean = False

    Public Sub detailspage()


    End Sub
    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        Weblib.LoginUser = request("MID") & ""
        Weblib.MerchantID = request("MEID") & ""

        Weblib.isStaff = weblib.bittoboolean(request("ist") & "")

        If isSkipMerchatID = False Then
            If (request("MEID") & "").trim = "" Then
                response.redirect("oops.aspx")
            End If
        End If

        Call checksecurity(request("MID") & "", request("SK") & "")

        If weblib.isstaff = True Then

            If loginpagestaff(Weblib.LoginUser) = False Then
                response.redirect("oops2.aspx")
            End If

        Else
            If loginpage(Weblib.LoginUser, Weblib.MerchantID) = False Then
                response.redirect("oops2.aspx")
            End If


        End If
        Call InitObjects()
    End Sub

    Public Function loginpagestaff(ByVal UserID As String) As Boolean

        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim dr As DataRow

        Try
            cmd.CommandText = "Select usr_email,usr_sysadmin,usr_code,usr_profile,usr_branch,usr_name,usr_firstscreen,usr_merchantid,isnull(usr_custbranchid,'') as usr_custbranchid,isnull(usr_custbranchnum,0) as usr_custbranchnum from secuserinfo where usr_loginid='" & userID.replace("'", "''") & "'  and isnull(usr_disable,0) = 0 "
            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")
            For Each dr In ds.Tables("datarecords").Rows
                WebLib.LoginUser = dr("usr_code").ToString.ToUpper
                WebLib.LoginUserName = dr("usr_name").ToString.ToUpper
                WebLib.StartupApp = dr("usr_firstscreen") & ""
                WebLib.ProfileID = dr("usr_profile").ToString.ToUpper
                WebLib.BranchID = dr("usr_branch").ToString.ToUpper

                Weblib.LoginIsFullAdmin = Weblib.BittoBoolean(dr("usr_sysadmin") & "")

                weblib.isstaff = True
                WebLib.CustBranchID = dr("usr_custbranchid").ToString.ToUpper
                WebLib.CustBranchNum = dr("usr_custbranchnum").ToString.ToUpper

                Try
                    WebLib.LoginUserEmail = dr("usr_email") & ""
                Catch ex As Exception

                End Try

                counter = counter + 1

                Exit For
            Next
            cn.Close()
            cmd.Dispose()
            cn.Dispose()


            If counter = 0 Then
                Return False

            Else
                Call WebLib.GetAppsByMerchantID(WebLib.Merchantid)
                Call WebLib.GetRightsByProfileID(WebLib.ProfileID)

                If (WebLib.CustBranchID & "").trim <> "" Then
                    Dim objbranch As New RuntimeCustomerBranch
                    objbranch.getInfo(WebLib.CustBranchNum, WebLib.CustNum)
                    weblib.custbranchname = objbranch.Description
                    objbranch = Nothing
                End If
                Return True
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function loginpage(ByVal UserID As String, ByVal MerchantID As String) As Boolean

        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim dr As DataRow

        Try
            cmd.CommandText = "Select usr_email,usr_merchantid,usr_code,usr_profile,usr_branch,usr_name,isnull(usr_custbranchid,'') as usr_custbranchid,isnull(usr_custbranchnum,0) as usr_custbranchnum from secuserinfo where usr_loginid='" & userID & "' and rtrim(isnull(usr_merchantid,'')) = '" & Merchantid & "'  and isnull(usr_disable,0) = 0 "
            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")
            For Each dr In ds.Tables("datarecords").Rows
                counter = counter + 1
                WebLib.LoginUser = dr("usr_code").ToString.ToUpper
                WebLib.ProfileID = dr("usr_profile").ToString.ToUpper
                WebLib.BranchID = dr("usr_branch").ToString.ToUpper
                WebLib.LoginUserName = dr("usr_name").ToString.ToUpper

                Try
                    WebLib.LoginUserEmail = dr("usr_email") & ""
                Catch ex As Exception
                    '    Return False
                End Try

                WebLib.CustCode = dr("usr_merchantid") & ""
                WebLib.Merchantid = dr("usr_merchantid") & ""
                weblib.isstaff = False
                WebLib.CustBranchID = dr("usr_custbranchid").ToString.ToUpper
                WebLib.CustBranchNum = dr("usr_custbranchnum").ToString.ToUpper

                Exit For
            Next
            cn.Close()
            cmd.Dispose()
            cn.Dispose()

            If counter = 0 Then
                Return False

                ' lblmessage.text = "Login Failed"

            Else
                Call initLogin()
                Call WebLib.GetRightsByProfileID(WebLib.ProfileID)
                If (WebLib.CustBranchID & "").trim <> "" Then
                    Dim objbranch As New RuntimeCustomerBranch
                    objbranch.getInfo(WebLib.CustBranchNum, WebLib.CustNum)
                    weblib.custbranchname = objbranch.Description
                    objbranch = Nothing
                End If

            End If

            Return True
        Catch ex As Exception

            Return False

            '            lblmessage.text = ex.Message
        End Try

    End Function


    Public Function checksecurity(ByVal UserID As String, ByVal SecurityKey As String) As Boolean

        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim dr As DataRow

        Try
            cmd.CommandText = "Select usr_merchantid from secuserinfo where usr_loginid='" & userID & "' and isnull(usr_disable,0) = 0 "
            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")
            For Each dr In ds.Tables("datarecords").Rows
                counter = counter + 1

                If (dr("usr_merchantid") & "").trim = "" Then
                    weblib.isstaff = True
                Else
                    weblib.isstaff = False

                End If
                Exit For
            Next
            cn.Close()
            cmd.Dispose()
            cn.Dispose()

        Catch ex As Exception

        End Try

    End Function



    Private Sub InitObjects()


        bid = Page.FindControl("bid")
        rid = Page.FindControl("rid")
        lblmessage = Page.FindControl("lblmessage")
        btnsave = Page.FindControl("SubmitButton")
        btnback = Page.FindControl("BackButton")

    End Sub
    Protected Sub InitLoad()


        rid.value = Request("ga") & ""
        bid.value = Request("ba") & ""

        If rid.value.trim <> "" Then
            _Mode = "M"
            Call ModifyMode()
        Else
            _Mode = "A"
            Call AddMode()
        End If

    End Sub
    Private Function initLogin() As Boolean
        If SOSettings.InitRuntimeObject = False Then
            '            Return False

            '            WebLib.ShowMessagePage(response, "Error in Sales Order Initialization", "main.aspx")
        Else
            WebLib.ProfileID = SOSettings.CustomerProfile
            WebLib.BranchID = SOSettings.CustomerBranch


        End If


        Dim obj As New RuntimeCustomer
        Call obj.getinfo(Weblib.MerchantID)
        weblib.CustNum = obj.CustNum
        weblib.CustName = obj.CustName
        obj = Nothing




    End Function

    Public Overridable Function LoadData() As Boolean

    End Function

    Protected Sub sendtoAPP(ByVal ActionCode As String)
        Page.ClientScript.RegisterStartupScript(Me.GetType(), "actioncode", "<script language='javascript'>location.href='" & ActionCode & "'</script>")
    End Sub

    Protected Sub gotoback()
        Response.Redirect(listingpage & "?ba=" & bid.value)
    End Sub
    Protected Sub backpage(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call gotoback()
    End Sub
    Protected Sub ModifyMode()
        If (modrights & "").trim <> "" Or (FullRights & "").trim <> "" Then
            If Weblib.hasrights(NmSpace, AppCode, modrights) = False And Weblib.hasrights(NmSpace, AppCode, FullRights) = False Then
                'WebLib.ShowMessagePage(response, "No Rights to Access this Feature", "supportdashboard.aspx")
                response.redirect("oops4.aspx")
            End If
        End If

        Call LoadData()
    End Sub
    Protected Sub AddMode()
        If (addrights & "").trim <> "" Then
            If Weblib.hasrights(NmSpace, AppCode, AddRights) = False And Weblib.hasrights(NmSpace, AppCode, FullRights) = False Then
                'WebLib.ShowMessagePage(response, "No Rights to Access this Feature", "supportdashboard.aspx")
                response.redirect("oops4.aspx")
            End If
        End If
        Call LoadData()
    End Sub
    Protected Sub DisableAll()

    End Sub

    Public Function savedata(ByVal insertfields As String, ByVal insertvalues As String, Optional ByVal _p_transaction As Boolean = False) As Boolean

        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()

        Try
            cn.Open()
            If rid.value = "" Then

                cmd.CommandText = "Insert into " & TableName & " (" & insertfields & ") Values (" & insertvalues & ")"
            Else

                If isnumeric(rid.value) = False Then
                    lblmessage.text = "Invalid ID to update"
                    Return False
                    Exit Function
                End If
                cmd.CommandText = "Update " & TableName & " set " & insertvalues & " where " & IDField & "=" & rid.value
            End If
            weblib.errortrap = cmd.commandtext
            cmd.Connection = cn
            cmd.ExecuteNonQuery()

            Return True
        Catch Err As Exception
            lblmessage.text = Err.message
            Return False
        Finally
            cn.Close()
            cmd = Nothing
        End Try

    End Function

End Class

