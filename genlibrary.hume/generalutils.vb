Imports Microsoft.VisualBasic
Imports System.Data.OleDB
Imports System.Data

Public Class clslogins
    Public ErrorMsg As String = ""
    Public connectionstring As String = System.Configuration.ConfigurationSettings.AppSettings("ConnStr")

    Public Function autologinstaff(ByVal ploginid As String) As Boolean
        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim dr As DataRow

        Try

            cmd.CommandText = "Select usr_sysadmin,usr_code,usr_profile,usr_branch,usr_name,usr_firstscreen,usr_merchantid,isnull(usr_custbranchid,'') as usr_custbranchid,isnull(usr_custbranchnum,0) as usr_custbranchnum,usr_email,usr_matrixlevel, usr_region, usr_isad from secuserinfo where usr_loginid='" & ploginid & "' and usr_filter='" & Weblib.FilterCode & "' and rtrim(isnull(usr_merchantid,'')) = '' and isnull(usr_disable,0) = 0 "
            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")
            For Each dr In ds.Tables("datarecords").Rows
                counter = counter + 1
                WebLib.LoginUserCompanySelected = "RGTECH"

                WebLib.LoginUser = dr("usr_code").ToString.ToUpper
                WebLib.LoginUserName = dr("usr_name").ToString.ToUpper
                WebLib.StartupApp = dr("usr_firstscreen") & ""
                WebLib.ProfileID = dr("usr_profile").ToString.ToUpper
                WebLib.BranchID = dr("usr_branch").ToString.ToUpper

                WebLib.Merchantid = ""

                Weblib.LoginIsFullAdmin = Weblib.BittoBoolean(dr("usr_sysadmin") & "")

                weblib.isstaff = True
                WebLib.CustBranchID = dr("usr_custbranchid").ToString.ToUpper
                WebLib.CustBranchNum = dr("usr_custbranchnum").ToString.ToUpper

                WebLib.LoginUserMatrixLevel = dr("usr_matrixlevel") & ""
                WebLib.LoginUserRegion = dr("usr_region") & ""

                WebLib.isAD = IIf(IsDBNull(dr("usr_isad")) Or WebLib.BitToBoolean(dr("usr_isad") & "") = False, False, True)
                WebLib.CustUnderLoginUserMatrixLevel = ""
                WebLib.LoginUserEmail = dr("usr_email") & ""

                Exit For
            Next
            cn.Close()
            cmd.Dispose()
            cn.Dispose()


            If counter = 0 Then
                ErrorMsg = "Login Failed"
                Return False
                Exit Function
            Else
                Call WebLib.GetAppsByMerchantID(WebLib.Merchantid)
                Call WebLib.GetRightsByProfileID(WebLib.ProfileID)

                webstats.trackstats("E")

                If (WebLib.CustBranchID & "").trim <> "" Then
                    Dim objbranch As New RuntimeCustomerBranch
                    objbranch.getInfo(WebLib.CustBranchNum, WebLib.CustNum)
                    weblib.custbranchname = objbranch.Description
                    objbranch = Nothing
                End If

                Return True
                Exit Function

            End If

        Catch ex As Exception
            Return False
            Exit Function

        End Try


    End Function


End Class
