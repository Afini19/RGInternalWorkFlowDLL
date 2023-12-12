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

Public Class detailspageopen

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

        Call InitObjects()
    End Sub

    Private Sub InitObjects()

    End Sub
    Protected Sub InitLoad()


    End Sub
    Public Overridable Function LoadData() As Boolean

    End Function

    Protected Sub gotoback()


    End Sub
    Protected Sub backpage(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub
    Protected Sub ModifyMode()

    End Sub
    Protected Sub AddMode()

    End Sub
    Protected Sub DisableAll()

    End Sub

    Public Function savedata(ByVal rid As String, ByVal insertfields As String, ByVal insertvalues As String, Optional ByVal _p_transaction As Boolean = False, Optional ByVal _p_addsql As String = "") As Boolean

        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()

        Try
            cn.Open()
            If rid = "" Then

                cmd.CommandText = "Insert into " & TableName & " (" & insertfields & ") Values (" & insertvalues & ")"
            Else

                If isnumeric(rid) = False Then
                    '                    lblmessage.text = "Invalid ID to update"
                    Return False
                    Exit Function
                End If
                cmd.CommandText = "Update " & TableName & " set " & insertvalues & " where " & IDField & "=" & rid
            End If

            If _p_addsql.trim <> "" Then
                cmd.commandtext = cmd.commandtext & ";" & _p_addsql
            End If


            Weblib.ErrorTrap = cmd.CommandText

            cmd.Connection = cn
            cmd.ExecuteNonQuery()

            Return True
        Catch Err As Exception
            Return False
        Finally
            cn.Close()
            cmd = Nothing
        End Try

    End Function

End Class

