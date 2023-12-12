Imports System
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.HttpServerUtility
Imports Microsoft.VisualBasic

Public Class globalfunction
    Inherits System.Web.UI.Page

    Public Shared Function convertDTForInsert(ByVal _p_datetime As DateTime)
        Return Format(_p_datetime.Year, "0000") & Format(_p_datetime.Month, "00") & Format(_p_datetime.Day, "00")
    End Function
    Public Shared Function convertDTStampForInsert(ByVal _p_datetime As DateTime)
        Return Format(_p_datetime.Year, "0000") & "-" & Format(_p_datetime.Month, "00") & "-" & Format(_p_datetime.Day, "00") & "T" & Format(_p_datetime.Hour, "00") & ":" & Format(_p_datetime.Minute, "00") & ":" & Format(_p_datetime.Second, "00")
    End Function
    Public Sub ValidatePage(ByVal response As Object)
        If (Session("EVENTCODE") & "").trim = "" Then
            response.redirect("selectevents.aspx")
        End If
    End Sub
    Public Sub ValidatePageSite(ByVal response As Object)


        If (Session("SELECTEDEVENT") & "").trim = "" Then
            response.redirect("default.aspx")
        End If
        Dim SettingsFile As String
        SettingsFile = System.AppDomain.CurrentDomain.BaseDirectory() & "EventData\" & Session("SELECTEDEVENT")
        Dim objset As New VIFEANDI_APP_INI
        Dim lClosingDate As String


        lClosingDate = objset.INIRead(SettingsFile, "EVENT DATA", "CLOSING DATE", "")
        objset = Nothing
        If microsoft.visualbasic.isdate(lClosingDate) = False Then
            lClosingDate = system.datetime.today
        End If
        If system.datetime.compare(system.datetime.parse(lClosingDate), system.datetime.today()) < 0 Then
            response.redirect("closing.aspx")
        End If
    End Sub
End Class
