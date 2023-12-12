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

Public Class stdpage

    Inherits System.Web.UI.Page
    Public touchscreenheight As String = ""
    Public connectionstring As String = System.Configuration.ConfigurationSettings.AppSettings("ConnStr")
    Public _pageindex As Long = 0
    Public _pagesize As Long = 20
    Public _searchkeystr = ""
    Public listingpage As String = ""
    Public _FormsName As String = ""
    Public columnscount As Integer = "1"
    Public TableName As String = ""
    Public DetailPage As String = ""
    Public IDPField As String = ""
    Public IDField As String = ""
    Public Orderby As String = ""
    Public AppIDField As String = ""
    Public MerchantIDField As String = ""
    Public FilterField As String = ""

    Public APPCode As String = ""
    Public AddRights As String = ""
    Public DelRights As String = ""
    Public ModRights As String = ""
    Public ViewRights As String = ""
    Public FullRights As String = ""
    Public NmSpace As String = ""

    Public Sub stdpage()


    End Sub
    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        If Weblib.LoginUser.Trim = "" Then
            response.redirect("~/login.aspx")
        End If


        Call InitObjects()
    End Sub
    Private Sub InitObjects()

    End Sub
    Protected Sub InitLoad()


    End Sub
End Class

