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

Public Class Imagesimages
    Public Shared Function GetImagebyStatus(ByVal Status As String) As String
        Select Case Status.ToLower
            Case "pending"
                Return WebLib.ClientURL("graphics/chops/pending300.png")
            Case "approved"
                Return WebLib.ClientURL("graphics/chops/approve300.png")

            Case "rejected"
                Return WebLib.ClientURL("graphics/chops/rejected300.png")
            Case "cancelled"
                Return WebLib.ClientURL("graphics/chops/cancel300.png")
            Case Else
                Return ""
        End Select
    End Function
End Class