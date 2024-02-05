Imports Microsoft.VisualBasic
Imports System.Data.OleDB
Imports System.Data

Public Class WebSearchCR
    Public ErrorMsg As String = ""
    Public connectionstring As String = System.Configuration.ConfigurationSettings.AppSettings("ConnStr")

    Public Shared Function CR_DocumentNo(ByVal docno As String)
        Return "cus_crno like '%" + docno.Replace("'", "''") + "%'"
    End Function

    Public Shared Function CR_RequestTitle(ByVal reqTitle As String)
        Return "cus_requestTitle like '%" + reqTitle.Replace("'", "''") + "%'"
    End Function

    Public Shared Function CR_Department(ByVal dept As String)
        Return "cus_department like '%" + dept.Replace("'", "''") + "%'"
    End Function

    Public Shared Function CR_Category(ByVal category1 As String)
        Return "cus_category like '%" + category1.Replace("'", "''") + "%'"
    End Function

    Public Shared Function CR_Module(ByVal module1 As String)
        Return "cus_module like '%" + module1.Replace("'", "''") + "%'"
    End Function

    Public Shared Function CR_Customer(ByVal customer1 As String)
        Return "cus_customer like '%" + customer1.Replace("'", "''") + "%'"
    End Function

    Public Shared Function CR_Priority(ByVal priority1 As String)
        Return "cus_priority like '%" + priority1.Replace("'", "''") + "%'"
    End Function

    Public Shared Function CR_TechnicalReq(ByVal techReq As String)
        Return "cus_technicalReq like '%" + techReq.Replace("'", "''") + "%'"
    End Function

    Public Shared Function CR_BusinessReq(ByVal businessReq As String)
        Return "cus_businessReq like '%" + businessReq.Replace("'", "''") + "%'"
    End Function

    Public Shared Function CR_Tags(ByVal tags1 As String)
        Return "cus_tags like '%" + tags1.Replace("'", "''") + "%'"
    End Function

    Public Shared Function CR_Status(ByVal status1 As String)
        Return "wst_status like '%" + status1.Replace("'", "''") + "%'"
    End Function

    Public Shared Function CR_PendingLevel(ByVal pendingLevel1 As String)
        Return "ApprovalLevelName like '%" + pendingLevel1.Replace("'", "''") + "%'"
    End Function

    Public Shared Function CR_PendingPerson(ByVal pendingPerson1 As String)
        Return "usr_name like '%" + pendingPerson1.Replace("'", "''") + "%'"
    End Function

    Public Shared Function CR_CreateDate(ByVal theDatainYYYYMMDD As String, ByVal theValue As String, ByVal OverrideOperator As String, Optional SearchAccuracy As Integer = 1)
        Return (" datediff(d,'" + theDatainYYYYMMDD + "',cus_createdt) " + OverrideOperator + " " + theValue.Replace("'", "''"))
    End Function

End Class
