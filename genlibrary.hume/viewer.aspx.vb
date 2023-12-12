Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Partial Public Class viewer_class
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        '        Dim lfile As String = HttpContext.Current.Server.URLEncode("http://www.humecementconnect.com.my/tempfiles/" & Request("file") & "")

        '       Response.Redirect("http://www.vifeandi.com/viewer/web/viewer.html?file=" & lfile)

        'Response.Redirect("http://www.vifeandi.com/viewer/web/viewer.html?file=http://www.humecementconnect.com.my/tempfiles/" & Request("file") & "")

        Response.Redirect("~/plugins/viewer/web/viewer.html?file=../../../tempfiles/" & Request("file") & "")
    End Sub
End Class

