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
Imports System.Collections
Imports System.Collections.Generic
Imports System.Xml

Public Class applistpage3

    Inherits System.Web.UI.Page
    Public connectionstring As String = System.Configuration.ConfigurationSettings.AppSettings("ConnStr")
    Public _pageindex As Long = 0
    Public _pagesize As Long = 20
    Public _searchkeystr As String = ""
    Public _searchfilter As String = ""
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
    Public pFieldNames As String = ""
    Public pJoinFields As String = ""
    Public _connection As String = ""
    Public _selectprefix As String = ""
    Public phSearchFields, phSearchDate As Object
    Public _SQLStatement As String = ""
    Public WithEvents rptPaging, rep As Repeater
    Public uc_from, uc_to, hrdate, lblmessage, lblCurrentPage, litSpacing, bid, rid, cmdprev, cmdnext, search_key1, btnadd As Object
    Public isSkipMerchatID As Boolean = False

    Public Sub applistpage3()


    End Sub
    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init

        If Weblib.MerchantID.trim = "" Then
            response.redirect("oops.aspx")
        End If
        If Weblib.LoginUser.trim = "" Then
            response.redirect("oops2.aspx")
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
    Public Sub AssignSearch()



        Dim lstr
        Dim counter As Integer = 0
        lstr = Microsoft.VisualBasic.Split(_searchkeystr, "|")
        If Microsoft.VisualBasic.UBound(lstr) = 0 Then

            Dim _rec As New Record_SearckKey
            If _rec.BreakKey(_searchkeystr) = True Then
                Dim chkbox As New CheckBox
                chkbox.ID = "chk" & _rec.KeyFieldName
                chkbox.Text = _rec.KeyCaption

                If _rec.KeyDataType <> "D" Then
                    chkbox.checked = True
                    phSearchFields.Controls.Add(chkbox)
                    phSearchFields.Controls.Add(New LiteralControl("<br>"))
                Else
                    phSearchDate.Controls.Add(chkbox)
                    phSearchDate.Controls.Add(New LiteralControl("<br>"))
                End If
            End If
            _rec = Nothing

        Else
            For counter = 0 To Microsoft.VisualBasic.UBound(lstr)

                Dim _rec As New Record_SearckKey
                If _rec.BreakKey(lstr(counter)) = True Then
                    Dim chkbox As New CheckBox
                    chkbox.ID = "chk" & _rec.KeyFieldName
                    chkbox.Text = _rec.KeyCaption


                    If _rec.KeyDataType <> "D" Then
                        chkbox.checked = True
                        phSearchFields.Controls.Add(chkbox)
                        phSearchFields.Controls.Add(New LiteralControl("<br>"))
                    Else
                        phSearchDate.Controls.Add(chkbox)
                        phSearchDate.Controls.Add(New LiteralControl("<br>"))

                    End If
                End If
                _rec = Nothing

            Next
        End If
        If phSearchDate.Controls.Count = 0 Then
            uc_from.visible = False
            uc_to.visible = False
            hrdate.visible = False
        End If

    End Sub

    Protected Sub cmdPrev_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        PageNumber = PageNumber - 1
        'loaddata()
        Call searchthedata()

    End Sub
    Protected Sub cmdNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        PageNumber = PageNumber + 1
        '        loaddata()
        Call searchthedata()

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


            _SQLStatement = sql

            Weblib.ErrorTrap = sql

            conn = New OdbcConnection(connectionString)
            conn.Open()
            comm = New OdbcCommand(sql, conn)

            ad.SelectCommand = comm
            ad.Fill(ds, "datarecords")


            If IsNumeric(_pageindex) = False Then
                _pageindex = 0
            End If

            Dim dt As DataTable = ds.Tables("datarecords")

            Dim dv As New DataView(dt)

            Dim pgitems As New PagedDataSource()
            pgitems.DataSource = dv
            pgitems.AllowPaging = True
            pgitems.PageSize = _pagesize
            pgitems.CurrentPageIndex = PageNumber
            If pgitems.PageCount > 1 Then
                rptPaging.Visible = True
                Dim pages As New ArrayList()
                For i As Integer = 0 To pgitems.PageCount - 1
                    pages.Add((i + 1).ToString())
                Next
                rptPaging.DataSource = pages
                rptPaging.DataBind()
            Else
                rptPaging.Visible = False
            End If

            lblCurrentPage.Text = "Page: " + (PageNumber + 1).ToString() & " of " & pgitems.PageCount.ToString()

            cmdPrev.Enabled = Not pgitems.IsFirstPage
            cmdNext.Enabled = Not pgitems.IsLastPage

            rep.DataSource = pgitems
            rep.DataBind()

            litSpacing.Text = ""
            For i As Integer = 0 To pgitems.PageSize - pgitems.Count
                litSpacing.Text = litSpacing.Text & "<br>"
            Next

            conn.Close()
            ' dr.Close()
            comm.Dispose()
            conn.Dispose()

        Catch ex As Exception
            '            lblmessage.text = ex.Message

        End Try

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
                    ltempwhere = " where " & ltempwhere & _searchfilter
                Else

                    If _searchfilter.trim <> "" Then
                        ltempwhere = " where " & _searchfilter
                    End If

                End If
                _p_searchkey = ltempwhere & Orderby

            End If


            If pFieldNames.trim = "" Then
                pFieldNames = "*"
            End If

            cn.Open()
            cmd.CommandText = "Select " & _selectprefix & " " & pFieldNames & " from " & TableName & " " & pJoinFields & " " & _p_searchkey

            _SQLStatement = cmd.CommandText
            Weblib.ErrorTrap = cmd.CommandText

            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")

            If IsNumeric(_pageindex) = False Then
                _pageindex = 0
            End If

            Dim dt As DataTable = ds.Tables("datarecords")

            Dim dv As New DataView(dt)

            Dim pgitems As New PagedDataSource()
            pgitems.DataSource = dv
            pgitems.AllowPaging = True
            pgitems.PageSize = _pagesize
            pgitems.CurrentPageIndex = PageNumber
            If pgitems.PageCount > 1 Then
                rptPaging.Visible = True
                Dim pages As New ArrayList()
                For i As Integer = 0 To pgitems.PageCount - 1
                    pages.Add((i + 1).ToString())
                Next
                rptPaging.DataSource = pages
                rptPaging.DataBind()
            Else
                rptPaging.Visible = False
            End If

            lblCurrentPage.Text = "Page: " + (PageNumber + 1).ToString() & " of " & pgitems.PageCount.ToString()

            cmdPrev.Enabled = Not pgitems.IsFirstPage
            cmdNext.Enabled = Not pgitems.IsLastPage

            rep.DataSource = pgitems
            rep.DataBind()

            litSpacing.Text = ""
            For i As Integer = 0 To pgitems.PageSize - pgitems.Count
                litSpacing.Text = litSpacing.Text & "<br>"
            Next

            cn.Dispose()
            cmd.Dispose()

        Catch ex As Exception
            '            lblmessage.text = ex.Message

        End Try

    End Sub
    Protected Sub DeleteRec(ByVal _p_ID As String)

        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()
        lblmessage.text = ""
        Dim _p_searchkey As String = " where " & IDField & "=" & _p_ID

        If isnumeric(_p_ID) = False Then
            lblmessage.text = "Delete Fail. Invalid ID"
            Exit Sub
        End If

        Dim ltempwhere As String = ""

        Try


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


            cn.Open()
            cmd.CommandText = "Delete from " & TableName & " " & _p_searchkey & ltempwhere
            cmd.Connection = cn
            cmd.ExecuteNonQuery()
            cn.Dispose()
            cmd.Dispose()

            Call loaddata()
        Catch ex As Exception
            lblmessage.text = ex.Message
        End Try



    End Sub
    Protected Sub rptPaging_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.RepeaterCommandEventArgs) Handles rptPaging.ItemCommand
        PageNumber = Convert.ToInt32(e.CommandArgument) - 1
        'loaddata()
        Call searchthedata()

    End Sub
    Public Sub addevent(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Response.Redirect("postpage.aspx?NextPage=" & DetailPage & "&ba=" & bid.value)
    End Sub
    Public Sub backpage(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call gotoback()
    End Sub
    Protected Sub gotoback()
        Response.Redirect(listingpage)
    End Sub
    Public Function GetSearchStatement(Optional ByVal ForceSearch As Boolean = False) As String
        Dim lSearchStr As String = ""
        Dim lSearchStr2 As String = ""

        Dim lstr
        Dim counter As Integer = 0
        Dim lKey1 As String = ""
        Dim lKey2 As String = ""
        Dim _l_EpicorDateFormat As String = "MM/dd/yyyy"

        lstr = Microsoft.VisualBasic.Split(_searchkeystr, "|")
        If Microsoft.VisualBasic.UBound(lstr) = 0 Then

            Dim _rec As New Record_SearckKey
            If _rec.BreakKey(_searchkeystr) = True Then
                If Request("chk" & _rec.KeyFieldName) = "on" Or ForceSearch = True Then
                    If _rec.KeyDataType = "D" Then
                        lKey1 = Microsoft.VisualBasic.Format(uc_from.DateValue, _l_EpicorDateFormat)
                        lKey2 = Microsoft.VisualBasic.Format(uc_to.DateValue, _l_EpicorDateFormat)

                    Else
                        lKey1 = search_key1.text
                        lKey2 = ""
                        lSearchStr = _rec.DefineSearchKey(_searchkeystr, lKey1, lKey2, _connection)

                    End If
                End If
            End If
            _rec = Nothing

        Else

            For counter = 0 To Microsoft.VisualBasic.UBound(lstr)

                Dim _rec As New Record_SearckKey
                If _rec.BreakKey(lstr(counter)) = True Then
                    If Request("chk" & _rec.KeyFieldName) = "on" Or ForceSearch = True Then

                        If _rec.KeyDataType = "D" Then
                            '                            lKey1 = Microsoft.VisualBasic.Format(uc_from.DateValue, _l_EpicorDateFormat)
                            '                           lKey2 = Microsoft.VisualBasic.Format(uc_to.DateValue, _l_EpicorDateFormat)
                            '                          If lSearchStr2.Trim = "" Then
                            'lSearchStr2 = "(" & _rec.DefineSearchKey(lstr(counter), lKey1, lKey2, _connection) & ")"
                            '                        Else
                            '                           lSearchStr2 = lSearchStr2 & " or (" & _rec.DefineSearchKey(lstr(counter), lKey1, lKey2, _connection) & ")"
                            '                      End If


                        Else
                            lKey1 = search_key1.text
                            lKey2 = ""
                            If lSearchStr.Trim = "" Then
                                lSearchStr = _rec.DefineSearchKey(lstr(counter), lKey1, lKey2)
                            Else
                                lSearchStr = lSearchStr & " or " & _rec.DefineSearchKey(lstr(counter), lKey1, lKey2)
                            End If

                        End If

                    Else

                    End If
                End If
                _rec = Nothing

            Next
        End If

        If lsearchstr.trim <> "" Then
            lsearchstr = "(" & lsearchstr & ") "
        End If

        If lsearchstr2.trim <> "" Then
            lsearchstr = lsearchstr & " and (" & lsearchstr2 & ")"
        End If

        Return lSearchStr

    End Function
    Private Sub searchthedata()
        Dim lSearchStr As String = GetSearchStatement() & ""
        Call loaddata(lSearchStr)
    End Sub
    Public Sub SearchStr(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call searchfromAPP()
    End Sub
    Public Sub searchfromAPP()
        PageNumber = 0
        Call searchthedata()
    End Sub
    Sub Grid_Change(ByVal sender As Object, ByVal e As DataGridPageChangedEventArgs)
        _pageindex = e.NewPageIndex
        '        Call loaddata()
        Call searchthedata()

    End Sub
    Protected Sub sendtoAPP(ByVal ActionCode As String)
        Page.ClientScript.RegisterStartupScript(Me.GetType(), "actioncode", "<script language='javascript'>location.href='" & ActionCode & "'</script>")
    End Sub

    Protected Sub InitLoad()
        If Page.IsPostBack = False Then


            If Weblib.hasrightsApp(AppCode) = False Then
                Try
                    Weblib.ShowMessagePage(response, "No rights to use this application sub module", "main.aspx")
                Catch ex As Exception

                End Try
            End If


            If (modrights & "").trim <> "" Or (FullRights & "").trim <> "" Or (ViewRights & "").trim <> "" Or (DelRights & "").trim <> "" Or (addRights & "").trim <> "" Then
                If Weblib.hasrights(NmSpace, AppCode, modrights) = False And Weblib.hasrights(NmSpace, AppCode, FullRights) = False And Weblib.hasrights(NmSpace, AppCode, AddRights) = False And Weblib.hasrights(NmSpace, AppCode, DelRights) = False And Weblib.hasrights(NmSpace, AppCode, ViewRights) = False Then
                    WebLib.ShowMessagePage(response, "No Rights to Access this Feature", "main.aspx")
                End If
            End If


            If Page.IsPostBack = False Then
                rid.value = Request("ga") & ""
                bid.value = Request("ba") & ""
            End If
            If IDPField.trim <> "" Then
                If bid.value.trim = "" Then
                    Call gotoback()
                End If
            End If
            Call loaddata()

            If Weblib.hasrights(NmSpace, AppCode, AddRights) = False And Weblib.hasrights(NmSpace, AppCode, FullRights) = False Then
                Try
                    btnAdd.Visible = False
                Catch ex As Exception

                End Try
            End If

        End If

        Call AssignSearch()
        lblmessage.text = ""


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

