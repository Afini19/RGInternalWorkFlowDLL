Imports System
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Configuration
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.HtmlControls
Imports System.Text
Imports System.Collections.Generic
Imports System.Xml
Imports System.Text.RegularExpressions

Public Class RuntimePersonalise

    Public Shared Property LastNamespace()
        Get
            If (HttpContext.Current.Session("LastNamespace") & "" <> "") Then
                Return HttpContext.Current.Session("LastNamespace")
            Else
                Return ""
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("LastNamespace") = value

        End Set

    End Property
    Public Shared Property LastQuery()
        Get
            If (HttpContext.Current.Session("LastQuery") & "" <> "") Then
                Return HttpContext.Current.Session("LastQuery")
            Else
                Return ""
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("LastQuery") = value

        End Set
    End Property
    Public Shared Property LastParameter1()
        Get
            If (HttpContext.Current.Session("LastParameter1") & "" <> "") Then
                Return HttpContext.Current.Session("LastParameter1")
            Else
                Return ""
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("LastParameter1") = value

        End Set
    End Property
    Public Shared Property LastParameter2()
        Get
            If (HttpContext.Current.Session("LastParameter2") & "" <> "") Then
                Return HttpContext.Current.Session("LastParameter2")
            Else
                Return ""
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("LastParameter2") = value

        End Set
    End Property

    Public Shared Function CanQuery(ByVal pNameSpace As String) As Boolean
        If pNamespace.tolower = LastNAmeSpace.tolower And (LastQuery.Tostring.trim <> "" Or LastParameter1.Tostring.trim <> "" Or LastParameter2.Tostring.trim <> "") Then
            Return True
        Else
            LastNameSpace = ""
            LastParameter1 = ""
            LastParameter2 = ""
            LastQuery = ""
            Return False
        End If
    End Function
End Class