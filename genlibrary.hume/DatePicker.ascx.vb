Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Partial Public Class datepicker_class
    Inherits System.Web.UI.UserControl
    Protected _Width As String = "150px"
    Protected _Height As String = "20px"
    Protected _IsDateOfBirth As Boolean = False
    Protected _yearrange As String = ""
    Protected _format As String = "dd/mm/yy"
    Protected _maxdate As Object = Nothing
    Public Placeholder As String = ""
    Public TextBox1, Literal1 As Object

    Public Property Textdmy() As String
        Get
            Return Weblib.Formatthedate(Value)
        End Get
        Set(ByVal value As String)
            TextBox1.text = value
        End Set
    End Property
    Public Property Value() As String
        Get
            Return TextBox1.text
        End Get
        Set(ByVal value As String)
            TextBox1.text = value
        End Set
    End Property
    Public Property ValidationGroup() As String
        Get
            Return TextBox1.ValidationGroup
        End Get
        Set(ByVal value As String)
            TextBox1.ValidationGroup = value
        End Set
    End Property
    Public Property Enabled() As Boolean
        Get
            Return TextBox1.Enabled
        End Get
        Set(ByVal value As Boolean)
            TextBox1.Enabled = value
        End Set
    End Property

    Public Property DateValue() As Date
        Get
            Dim hasError As Boolean = False
            Dim lyear As Integer
            Try
                lyear = CInt(Microsoft.VisualBasic.Mid(Value, 7, 4))

            Catch ex As Exception
                haserror = True
                lyear = 1991
            End Try
            Dim lmonth As Integer
            Try
                lmonth = CInt(Microsoft.VisualBasic.Mid(Value, 1, 2))

            Catch ex As Exception
                haserror = True
                lmonth = 1
            End Try
            Dim lday As Integer
            Try
                lday = CInt(Microsoft.VisualBasic.Mid(Value, 4, 2))

            Catch ex As Exception
                haserror = True
                lday = 1
            End Try

            Try

                If haserror = True Then
                    Return New DateTime(1991, 1, 1)
                Else
                    Dim justDate As DateTime = New DateTime(lyear, lmonth, lday)
                    Return justDate
                End If

            Catch ex As Exception
                Return New DateTime(1991, 1, 1)
            End Try
        End Get
        Set(ByVal value As Date)

        End Set
    End Property

    Public Property IsDateOfBirth() As Boolean
        Get
            Return _IsDateOfBirth
        End Get
        Set(ByVal value As Boolean)
            _IsDateOfBirth = value
            If value = True Then
                _yearrange = ", yearRange: '1900:" & DateTime.Now.Year & "'"
            Else
                _yearrange = ""
            End If

        End Set
    End Property
    Public Property Max() As Object
        Get
            If isnothing(_maxdate) = True Then
                Return Nothing
            Else
                Return _maxdate
            End If
        End Get
        Set(ByVal value As Object)
            _maxdate = value
        End Set
    End Property

    Public Property cssclass() As String
        Get
            Return TextBox1.cssclass
        End Get
        Set(ByVal value As String)
            TextBox1.cssclass = value
        End Set
    End Property
    Public Property Width() As String
        Get
            Return _Width
        End Get
        Set(ByVal value As String)
            _Width = value
            TextBox1.Style.Add("width", _Width)
        End Set
    End Property
    Public Property Height() As String
        Get
            Return _Height
        End Get
        Set(ByVal value As String)
            _Height = value
            TextBox1.Style.Add("height", _Height)
        End Set
    End Property
    Protected Overrides Sub OnPreRender(ByVal e As System.EventArgs)
        MyBase.OnPreRender(e)


        Dim ltemp As String = ""
        Dim ldaterange As String = ""
        '        ldaterange = ", minDate: new Date(" + Min.Value.Year + ", " + Min.Value.Month + " - 1, " + Min.Value.Day + ")"

        If isnothing(max) = False Then
            ldaterange = ", maxDate: new Date(" & Max.Year & ", " & Max.Month & " - 1, " & Max.Day & ")"
        End If

        ltemp = ltemp & "$(function () {" & environment.NEwLine
        '        ltemp = ltemp & "$('#" & TextBox1.ClientID + "').datepicker( {showButtonPanel: true, changeMonth: true, changeYear: true, showOtherMonths: true, selectOtherMonths: true, numberOfMonths: 2" & ldateRange & _yearrange & ", onClose: function() {this.focus();} } );" & environment.NEwLine
        ltemp = ltemp & "$('#" & TextBox1.ClientID + "').datepicker( {showButtonPanel: true, changeMonth: true, changeYear: true, showOtherMonths: true, selectOtherMonths: true, numberOfMonths: 1" & ldateRange & _yearrange & "} );" & environment.NEwLine
        ltemp = ltemp & "$('#" & TextBox1.ClientID + "').datepicker( ""option"", ""dateFormat"",""" & _format & """);" & environment.NEwLine


        ltemp = ltemp & "$('#" & TextBox1.ClientID + "').attr('placeholder','" & Placeholder & "');" & environment.NEwLine

        ltemp = ltemp & "});   " & environment.NEwLine


        Literal1.text = "<script>" & environment.NEwLine & ltemp & environment.NEwLine & "</script>"
    End Sub
    Protected Sub Page_LoadComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If TextBox1.text.length = 10 Then
            TextBox1.text = TextBox1.text.Substring(3, 2) + "/" + TextBox1.text.Substring(0, 2) + "/" + TextBox1.text.Substring(6, 4)

        End If
    End Sub
End Class

