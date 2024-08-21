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

Public Class detailspageV2

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
    Public bid, rid, lblmessage, btnsave As Object
    Public _Mode As String = ""
    Public fUserID As String = ""
    Public APPCode As String = ""
    Public AddRights As String = ""
    Public DelRights As String = ""
    Public ModRights As String = ""
    Public ViewRights As String = ""
    Public FullRights As String = ""
    Public NmSpace As String = ""

    Public Sub detailspage()


    End Sub
    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        If Weblib.LoginUser.Trim = "" Then
            response.redirect("~/loginstaff.aspx")
        End If


        Call InitObjects()
    End Sub
    Public Sub LockScreenAfterClick(ByVal wc As WebControl, ByVal message As String)
        AddLockScreenScript()
        wc.Attributes("onclick") = String.Format("skm_LockScreen('{0}');", message.Replace("'", "\'"))
    End Sub

    Public Sub LockScreenAfterDDLChange(ByVal ddl As DropDownList, ByVal message As String)
        AddLockScreenScript()
        ddl.Attributes("onchange") = String.Format("skm_LockScreen('{0}');", message.Replace("'", "\'"))
    End Sub
    Private Sub AddLockScreenScript()
        'Add the JavaScript and <div> elements for freezing the screen
        If Not ClientScript.IsClientScriptIncludeRegistered("skm_LockScreen") Then
            'Register the JavaScript Include
            ClientScript.RegisterClientScriptInclude("skm_LockScreen", Page.ResolveUrl("~/Scripts/LockScreen.js?version=1.0"))

            'Add the <div> elements
            ClientScript.RegisterClientScriptBlock(Me.GetType(), "skm_LockScreen_divs", _
                                                    "<div id=""skm_LockBackground"" class=""LockOff""></div><div id=""skm_LockPane"" class=""LockOff""><div id=""skm_LockPaneText"">&nbsp;</div></div>", _
                                                    False)
        End If
    End Sub
    Private Sub InitObjects()
        bid = Page.FindControl("bid")
        rid = Page.FindControl("rid")
        lblmessage = Page.FindControl("lblmessage")
        btnsave = Page.FindControl("SubmitButton")

    End Sub
    Protected Sub InitLoad()



        fUserID = WebLib.LoginUser
        If Page.IsPostBack = False Then

            If Weblib.hasrightsApp(AppCode) = False Then
                Try
                    Weblib.ShowMessagePage(response, "No rights to use this application sub module", "main.aspx")
                Catch ex As Exception

                End Try
            End If


            rid.value = Request("ga") & ""
            bid.value = Request("ba") & ""

            If rid.value.trim <> "" Then
                _Mode = "M"
                Call ModifyMode()
            Else
                _Mode = "A"
                Call AddMode()
            End If


            If addrights.trim <> "" Or modrights.trim <> "" Or fullrights.trim <> "" Then

                If Weblib.hasrights(NmSpace, AppCode, AddRights) = False And Weblib.hasrights(NmSpace, AppCode, ModRights) = False And Weblib.hasrights(NmSpace, AppCode, FullRights) = False Then
                    Try
                        If listingpage.trim = "" Then
                            Weblib.ShowMessagePage(response, "No rights to use this application sub module", "main.aspx")

                        Else
                            btnsave.Visible = False

                        End If
                    Catch ex As Exception

                    End Try
                End If
            End If


            Try
                If (request("msg") & "").trim <> "" Then
                    lblmessage.text = Weblib.getAlertMessageStyle(request("msg") & "")
                End If

            Catch ex As Exception

            End Try

        End If

    End Sub
    Public Overridable Function LoadData() As Boolean

    End Function


    Protected Sub gotoback()
        Response.Redirect(listingpage & "?ba=" & bid.value)
    End Sub
    Protected Sub backpage(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call gotoback()
    End Sub
    Protected Sub ModifyMode()
        If (modrights & "").trim <> "" Or (FullRights & "").trim <> "" Then
            If Weblib.hasrights(NmSpace, AppCode, modrights) = False And Weblib.hasrights(NmSpace, AppCode, FullRights) = False Then
                WebLib.ShowMessagePage(response, "No Rights to Access this Feature", "supportdashboard.aspx")
            End If
        End If
        Call LoadData()
    End Sub
    Protected Sub AddMode()
        Call LoadData()
    End Sub
    Protected Sub DisableAll()

    End Sub

    Public Function savedata(ByVal insertfields As String, ByVal insertvalues As String, Optional ByVal _p_transaction As Boolean = False, Optional ByVal _p_addsql As String = "") As Boolean

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

            If _p_addsql.trim <> "" Then
                cmd.commandtext = cmd.commandtext & ";" & _p_addsql
            End If

            LogtheAudit(cmd.CommandText)
            Weblib.ErrorTrap = cmd.CommandText

            cmd.Connection = cn
            cmd.ExecuteNonQuery()

            Return True
        Catch Err As Exception
            lblmessage.text = Err.Message
            Return False
        Finally
            cn.Close()
            cmd = Nothing
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

