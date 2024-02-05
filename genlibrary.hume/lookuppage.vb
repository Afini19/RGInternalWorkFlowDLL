Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System
Imports System.Configuration
Imports System.Collections.Generic
Imports System.Xml

Public Class lookuppage

    Inherits System.Web.UI.Page
    Public connectionstring As String = System.Configuration.ConfigurationSettings.AppSettings("ConnStr")
    Public _pageindex As Long = 0
    Public _pagesize As Long = 10
    Public _searchkeystr = ""
    Public _searchfilter As String = ""

    Public listingpage As String = ""
    Public _FormsName As String = ""
    Public columnscount As Integer = "7"
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
    Public pFieldNames As String = ""
    Public pJoinFields As String = ""
    Public _connection As String = ""
    Public _returnfield As String = ""
    Public _returnfield2 As String = ""
    Public _returnfield3 As String = ""
    Public _returnfield4 As String = ""
    Public _returnfield5 As String = ""
    Public _returnfield6 As String = ""
    Public _returnfield7 As String = ""
    Public _returnfield8 As String = ""
    Public _returnfield9 As String = ""
    Public _returnfield10 As String = ""

    Public _previewfield As String = ""

    Public phSearchFields, phSearchDate As Object
    Public WithEvents rptPaging, rep As Repeater
    Public uc_from, uc_to, hrdate, lblmessage, lblCurrentPage, litSpacing, bid, rid, cmdprev, cmdnext, search_key1, btnadd As Object
    Public Sub listpage()


    End Sub
    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init

        If Weblib.LoginUser.Trim = "" Then
            response.redirect("loginstaff.aspx")
        End If

        Call InitObjects()
    End Sub

    Public Property PageNumber() As Integer
        Get
            If ViewState("PageNumber") IsNot Nothing Then
                Return Convert.ToInt32(ViewState("PageNumber"))
            Else
                Return 0
            End If
        End Get
        Set(ByVal value As Integer)
            ViewState("PageNumber") = value
        End Set
    End Property
    Private Sub InitObjects()
        btnAdd = Page.FindControl("btnAdd")
        phSearchFields = Page.FindControl("phSearchFields")
        phSearchDate = Page.FindControl("phSearchDate")
        uc_from = Page.FindControl("uc_from")
        uc_to = Page.FindControl("uc_to")
        hrdate = Page.FindControl("hrdate")
        lblmessage = Page.FindControl("lblmessage")
        lblCurrentPage = Page.FindControl("lblCurrentPage")
        litSpacing = Page.FindControl("litSpacing")
        rptPaging = Page.FindControl("rptPaging")
        rep = Page.FindControl("rep")
        bid = Page.FindControl("bid")
        rid = Page.FindControl("rid")
        cmdprev = Page.FindControl("cmdprev")
        cmdnext = Page.FindControl("cmdnext")
        search_key1 = Page.FindControl("search_key1")
    End Sub

    Protected Sub cmdPrev_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        PageNumber = PageNumber - 1
        loaddata()
    End Sub
    Protected Sub cmdNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        PageNumber = PageNumber + 1
        loaddata()
    End Sub
    Public Sub loaddataODBC(Optional ByVal _p_searchkey As String = "")


        Dim conn As OdbcConnection
        Dim comm As OdbcCommand
        Dim dr As DataRow
        Dim counter As Integer = 0
        Dim ad As New Odbc.OdbcDataAdapter
        Dim ds As New DataSet()
        Dim connectionString As String
        Dim sql As String



        '        lblmessage.text = ""

        '        Try


        Dim ltempwhere As String = ""

        If IDPField.trim <> "" Then
            If isnumeric(bid.value) = False Then
                Exit Sub
            End If
            ltempwhere = ltempwhere & " and " & IDPField & "=" & bid.value
        End If
        If AppIDField.trim <> "" Then
            ltempwhere = ltempwhere & " and " & AppIDField & "='" & WebLib.ApplicationID & "'"
        End If
        If MerchantIDField.trim <> "" Then
            ltempwhere = ltempwhere & " and " & MerchantIDField & "='" & WebLib.MerchantID & "'"
        End If
        If FilterField.trim <> "" Then
            ltempwhere = ltempwhere & " and " & FilterField & "='" & WebLib.FilterCode & "'"
        End If

        If (_p_searchkey.trim <> "" Or ltempwhere.trim <> "") And _searchfilter.trim <> "" Then
            _searchfilter = " and " & _searchfilter & " "
        End If

        If _p_searchkey.Trim <> "" Then
            _p_searchkey = " where (" & _p_searchkey & ") " & ltempwhere & " " & _searchfilter & Orderby
        Else

            If Microsoft.visualbasic.left(ltempwhere, 4) = " and" Then  'thespace have to be there
                ltempwhere = Microsoft.visualbasic.right(ltempwhere, ltempwhere.length - 4)
            End If
            ltempwhere = ltempwhere.trim

            If ltempwhere.trim <> "" Then
                ltempwhere = " where " & ltempwhere
            End If

            If ltempwhere.trim = "" And _searchfilter.trim <> "" Then

                _p_searchkey = "where " & _searchfilter & Orderby

            Else
                _p_searchkey = ltempwhere & _searchfilter & Orderby

            End If

        End If

        If pFieldNames.trim = "" Then
            pFieldNames = "*"
        End If



        connectionString = Weblib.connEpicor
        sql = "Select " & pFieldNames & " from " & TableName & " " & pJoinFields & " " & _p_searchkey
        weblib.errortrap = sql
        conn = New OdbcConnection(connectionString)
        conn.Open()
        comm = New OdbcCommand(sql, conn)

        ad.SelectCommand = comm
        ad.Fill(ds, "datarecords")


        If IsNumeric(_pageindex) = False Then
            _pageindex = 0
        End If

        Dim dt As DataTable = ds.Tables("datarecords")

        '            Dim dv As New DataView(dt)

        '            Dim pgitems As New PagedDataSource()
        '            pgitems.DataSource = dv
        '            pgitems.AllowPaging = True
        '            pgitems.PageSize = _pagesize
        '            pgitems.CurrentPageIndex = PageNumber
        '            If pgitems.PageCount > 1 Then
        ' rptPaging.Visible = True
        ' Dim pages As New ArrayList()
        ' For i As Integer = 0 To pgitems.PageCount - 1
        ' pages.Add((i + 1).ToString())
        ' Next
        ' rptPaging.DataSource = pages
        ' rptPaging.DataBind()
        ' Else
        ' rptPaging.Visible = False
        ' End If

        '            lblCurrentPage.Text = "Page: " + (PageNumber + 1).ToString() & " of " & pgitems.PageCount.ToString()

        '            cmdPrev.Enabled = Not pgitems.IsFirstPage
        '           cmdNext.Enabled = Not pgitems.IsLastPage

        rep.DataSource = dt  'pgitems
        rep.DataBind()

        litSpacing.Text = ""
        '            For i As Integer = 0 To pgitems.PageSize - pgitems.Count
        'litSpacing.Text = litSpacing.Text & "<br>"
        'Next

        conn.Close()
        ' dr.Close()
        comm.Dispose()
        conn.Dispose()

        '        Catch ex As Exception
        '        lblmessage.text = ex.Message

        '      End Try

    End Sub

    Public Sub loaddata(Optional ByVal _p_searchkey As String = "")


        Select Case _connection

            Case "Epicor"
                Call loaddataODBC(_p_searchkey)
                Exit Sub
            Case "FMS"
                connectionstring = WebLib.ConnFMS
        End Select


        Dim cn As New OleDbConnection(connectionstring)

        Dim cmd As New OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim dr As DataRow
        lblmessage.text = ""

        Try


            Dim ltempwhere As String = ""

            If IDPField.trim <> "" Then
                If isnumeric(bid.value) = False Then
                    Exit Sub
                End If
                ltempwhere = ltempwhere & " and " & IDPField & "=" & bid.value
            End If
            If AppIDField.trim <> "" Then
                ltempwhere = ltempwhere & " and " & AppIDField & "='" & WebLib.ApplicationID & "'"
            End If
            If MerchantIDField.trim <> "" Then
                ltempwhere = ltempwhere & " and " & MerchantIDField & "='" & WebLib.MerchantID & "'"
            End If
            If FilterField.trim <> "" Then
                ltempwhere = ltempwhere & " and " & FilterField & "='" & WebLib.FilterCode & "'"
            End If

            If (_p_searchkey.trim <> "" Or ltempwhere.trim <> "") And _searchfilter.trim <> "" Then
                _searchfilter = " and " & _searchfilter & " "
            End If

            If _p_searchkey.Trim <> "" Then
                _p_searchkey = " where (" & _p_searchkey & ") " & ltempwhere & " " & _searchfilter & Orderby
            Else

                If Microsoft.visualbasic.left(ltempwhere, 4) = " and" Then  'thespace have to be there
                    ltempwhere = Microsoft.visualbasic.right(ltempwhere, ltempwhere.length - 4)
                End If
                ltempwhere = ltempwhere.trim

                If ltempwhere.trim <> "" Then
                    ltempwhere = " where " & ltempwhere
                End If
                '_p_searchkey = ltempwhere & Orderby
                '                _p_searchkey = ltempwhere & _searchfilter & Orderby

                If ltempwhere.trim = "" And _searchfilter.trim <> "" Then

                    _p_searchkey = "where " & _searchfilter & Orderby

                Else
                    _p_searchkey = ltempwhere & _searchfilter & Orderby

                End If


            End If

            If pFieldNames.trim = "" Then
                pFieldNames = "*"
            End If

            cn.Open()
            cmd.CommandText = "Select " & pFieldNames & " from " & TableName & " " & pJoinFields & " " & _p_searchkey

            weblib.errortrap = cmd.CommandText

            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")

            If IsNumeric(_pageindex) = False Then
                _pageindex = 0
            End If

            Dim dt As DataTable = ds.Tables("datarecords")

            '            Dim dv As New DataView(dt)

            '           Dim pgitems As New PagedDataSource()
            '          pgitems.DataSource = dv
            '          pgitems.AllowPaging = True
            '          pgitems.PageSize = _pagesize
            '          pgitems.CurrentPageIndex = PageNumber
            '          If pgitems.PageCount > 1 Then
            ' rptPaging.Visible = True
            ' Dim pages As New ArrayList()
            ' For i As Integer = 0 To pgitems.PageCount - 1
            ' pages.Add((i + 1).ToString())
            ' Next
            ' rptPaging.DataSource = pages
            ' rptPaging.DataBind()
            'Else
            '    rptPaging.Visible = False
            'End If

            'lblCurrentPage.Text = "Page: " + (PageNumber + 1).ToString() & " of " & pgitems.PageCount.ToString()

            'cmdPrev.Enabled = Not pgitems.IsFirstPage
            'cmdNext.Enabled = Not pgitems.IsLastPage

            rep.DataSource = dt  'pgitems
            rep.DataBind()

            '            litSpacing.Text = ""
            '            For i As Integer = 0 To pgitems.PageSize - pgitems.Count
            ' litSpacing.Text = litSpacing.Text & "<br>"
            ' Next

            cn.Dispose()
            cmd.Dispose()

        Catch ex As Exception
            '            lblmessage.text = ex.Message

        End Try

    End Sub
    Protected Sub rptPaging_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.RepeaterCommandEventArgs) Handles rptPaging.ItemCommand
        PageNumber = Convert.ToInt32(e.CommandArgument) - 1
        loaddata()
    End Sub

    Sub Grid_Change(ByVal sender As Object, ByVal e As DataGridPageChangedEventArgs)
        _pageindex = e.NewPageIndex
        Call loaddata()
    End Sub
    Protected Sub InitLoad()
        If Page.IsPostBack = False Then

            If Weblib.hasrightsApp(AppCode) = False Then
                Try
                    Weblib.ShowMessagePage(response, "No rights to use this application sub module", "main.aspx")
                Catch ex As Exception

                End Try
            End If


            rid.value = Request("ga") & ""
            bid.value = Request("ba") & ""

            If IDPField.trim <> "" Then
                If bid.value.trim = "" Then

                End If
            End If
            Call loaddata()

        End If

        lblmessage.text = ""

    End Sub
End Class

