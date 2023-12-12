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

Namespace Vifeandi.Custom
    Public Class Redirect
        Public connectionstring As String = System.Configuration.ConfigurationSettings.AppSettings("ConnStr")

        Public Function encodeIndex(ByVal index As Integer) As String
            Dim v1, v2 As Integer
            Dim s1, s2 As String
            Dim s As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
            v1 = Int((33 * Rnd()) + 1)
            v2 = Int((88 * Rnd()) + 1)
            s1 = s.Substring(((v1 Mod 35) + index), 1)
            s2 = s.Substring(((v2 Mod 35) + index), 1)

            Return s1 & s2
        End Function

        Public Function web_redirect(ByVal URL As String, ByVal ValidDays As Integer, ByVal maxclick As Integer, ByVal userid As String, ByVal intrefid As String, ByVal strrefid2 As String, ByVal strrefid3 As String, Optional ByVal index As Integer = 0) As String
            Dim ltempstring = ""
            Try

                Dim strConnection As String = connectionstring
                Dim cnHQ As New OleDb.OleDbConnection
                Dim cmdHQ As New OleDb.OleDbCommand

                Dim lsql As String
                Dim lCreateDate As DateTime
                lCreateDate = DateTime.Today
                Dim lExpiry As String
                Dim clicks As String
                Dim lcode As String

                Dim oo As New vifeandi.general
                lcode = oo.GetUniqueCode(4) & encodeIndex(index)


                If ValidDays = -1 Then
                    lExpiry = "Null"
                Else
                    lExpiry = "'" & oo.convertDTStampForInsert(lCreateDate.Date.AddDays(ValidDays)) & "'"
                End If
                If maxclick = -1 Then
                    clicks = "Null"
                Else
                    clicks = maxclick
                End If

                oo = Nothing

                strConnection = connectionstring

                Try
                    cnHQ = New OleDb.OleDbConnection(strConnection)
                    cnHQ.Open()
                    cmdHQ.Connection = cnHQ

                    If isnumeric(intrefid) = False Then
                        intrefid = 0
                    End If

                    lsql = "Insert into [Redirect]([red_code],[red_redirectto],[red_expiry],[red_time],red_user,red_refID,red_refID2,red_refID3) Values ("
                    lsql = lsql & "'" & lcode & "','" & URL & "'," & lExpiry & "," & clicks & ",'" & userid & "'," & intrefid & ",'" & strrefid2 & "','" & strrefid3 & "')"

                    cmdHQ.CommandText = lsql
                    cmdHQ.ExecuteNonQuery()
                    cmdHQ.Dispose()
                    cnHQ.Close()
                    cnHQ.Dispose()
                    Return lcode

                Catch ex As Exception
                    Return ""
                Finally


                End Try


            Catch ex As Exception
                Return ""


            End Try


        End Function


    End Class
End Namespace
