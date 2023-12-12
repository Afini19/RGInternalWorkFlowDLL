Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Partial Public Class calculator_class
    Inherits System.Web.UI.UserControl
    Private _forPassword As Boolean = False
    Public TextBox1 As Object

    Public Property ForPassword() As Boolean
        Get
            Return _forPassword
        End Get
        Set(ByVal value As Boolean)
            _forPassword = value
            If _forPassword = True Then

                TextBox1.TextMode = TextBoxMode.Password

            Else
                TextBox1.TextMode = TextBoxMode.SingleLine


            End If
        End Set
    End Property
    Public Property Text() As String
        Get
            Return TextBox1.Text
        End Get
        Set(ByVal value As String)

            TextBox1.Text = value
        End Set
    End Property
End Class

