Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Partial Public Class workflowbar_class
    Inherits System.Web.UI.UserControl

    Public Event OnApproveEvent As System.EventHandler
    Public Event OnRejectEvent As System.EventHandler
    Public Event OnCancelEvent As System.EventHandler

    Private connectionstring As String = System.Configuration.ConfigurationSettings.AppSettings("ConnStr")
    Public parentobj As Object
    Public wstatus, lblmessage, wlevel, wlevelA, wucode, btnwapprove, btnwreject, btnwcancel, actype, artype, aatype, wwid, cus_wrefno, btntracking, btnattachments, litstatus, litdetails, wr, wrm, litapprovalperson, litattachments, litAudit, wlevelamt, wlevelamtend, wlevelamtenabled, wversion, btnwresend As Object
    Public workflownamespace As String = ""
    Public custommode As Boolean = False
    Public overridemode As Boolean = False
    Public attachmentbycreatoronly As Boolean = False
    Public isapproved As Boolean = False
    Public canattach As Boolean = False
    Dim _l_currentamt As String = 0
    Public ErrorMsg As String = ""
    Public StrictValidationGroup As Boolean = False
    Public Property uid() As String
        Get
            Return wucode.value
        End Get
        Set(ByVal value As String)
            wucode.value = value

            Try
                btnwresend.visible = True
            Catch ex As Exception
                btnwresend.visible = False
            End Try
            Call loaddata()
        End Set
    End Property
    Public Property DocumentAmt() As String
        Get
            If isnumeric(_l_currentamt) = False Then
                _l_currentamt = 0
            End If
            Return _l_currentamt
        End Get
        Set(ByVal value As String)
            If isnumeric(value) = False Then
                _l_currentamt = 0
            Else
                _l_currentamt = value
            End If
        End Set
    End Property
    Public Property WorkflowEnded() As Boolean
        Get
            If wstatus.Value.ToLower = "pending" Or wstatus.Value.ToLower.Trim = "" Then
                Return False
            Else
                isapproved = True
                Return True
            End If

        End Get
        Set(ByVal value As Boolean)

        End Set
    End Property
    Public Function LoadData() As Boolean
        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim dr As DataRow
        Dim templi As ListItem



        litdetails.Text = "<font color=""red"">Pending Document Creation<font>"


        If wucode.Value.Trim = "" Then

            Call setstatus("")
            Exit Function
        End If

        If (wr.Value & "").Trim = "" Then
            Call GetRights()
        End If


        Try
            cmd.CommandText = "Select * from workflowstatus where wst_ucode='" & wucode.Value & "'"
            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")
            For Each dr In ds.Tables("datarecords").Rows
                counter = counter + 1
                wwid.Value = dr("wst_workflowid") & ""
                wlevel.Value = dr("wst_level") & ""

                cus_wrefno.Value = dr("wst_refno") & ""
                wstatus.Value = dr("wst_status") & ""


                getLevelA()
                If (dr("wst_lastupdateon") & "").trim <> "" Then
                    litdetails.Text = "Current Approval Level : <b>" & dr("wst_level") & "</b><br>Last Action On :-<br>" & dr("wst_lastupdateon") & ""
                Else
                    litdetails.Text = "Current Approval Level : <b>" & dr("wst_level") & "</b>"
                End If
                Call setstatus(dr("wst_status") & "")

                Exit For
            Next
            cn.Close()
            cmd.Dispose()
            cn.Dispose()


            If StrictValidationGroup = True Then
                EnableDisable2(Page)
            Else
                EnableDisable(Page)
            End If


            LoadAudit()

        Catch ex As Exception
            lblMessage.Text = ex.Message
        End Try
    End Function
    Private Function getLevelA()
        Dim ooWorkflow As New clsWorkflow
        wlevelA.Value = ooWorkflow.GetLevelNobySequence(wwid.Value, wlevel.Value)
        ooWorkFlow = Nothing
    End Function
    Private Function getAuditSQL(ByVal pType As String, ByVal pRefNo As String, ByVal pDescription As String, ByVal pucode As String) As String

        Dim ooWorkFlow As New clsWorkFlow
        Return ooWorkFlow.getAuditSQL(wlevel.value, pType, pRefNo, pDescription, pucode)
        ooWorkflow = Nothing
        Exit Function

        '        Dim pLevel As String
        '       pLevel = wlevel.Value

        '        If IsNumeric(pLevel) = False Then
        'pLevel = "NULL"
        ' End If

        '        Return "Insert into workflowaudit (wfa_code,wfa_refno,wfa_description,wfa_ucode,wfa_createon,wfa_createby,wfa_merchantid,wfa_filtercode,wfa_level) Values " & _
        '                       "('" & pType & "','" & pRefNo & "','" & pDescription & "','" & pucode & "',getdate(),'" & WebLib.LoginUser & "','" & WebLib.MerchantID & "','" & WebLib.FilterCode & "'," & pLevel & ")"

    End Function
    Public Function GetWorkFlowSaveSQL(ByVal WorkflowID As String, ByVal UniqueCode As String, ByVal isCreate As Boolean, ByVal RefNo As String, ByVal Modulecode As String, ByVal DocumentTitle As String, Optional ByVal Param1 As String = "", Optional ByVal Param2 As String = "", Optional ByVal Param3 As String = "", Optional ByVal Param4 As String = "") As String
        Dim ltempsql As String
        If isCreate = True Then
            If IsNumeric(WorkflowID) = False Then
                Return ""
                Exit Function
            End If
            If Modulecode.Trim = "" Then
                Return ""
                Exit Function
            End If
            If UniqueCode.Trim = "" Then
                Return ""
                Exit Function
            End If

            'ltempsql = "Insert into workflowstatus (wst_workflowid,wst_ucode,wst_refno,wst_module,wst_status,wst_level,wst_merchantid,wst_filtercode,wst_createon,wst_createby,wst_subject) Values " & _
            '           "(" & WorkflowID & ",'" & UniqueCode & "','" & RefNo.Replace("'", "''") & "','" & Modulecode & "','Pending',1,'" & WebLib.MerchantID & "','" & WebLib.FilterCode & "',getdate(),'" & WebLib.LoginUser & "','" & DocumentTitle.ToString.Replace("'", "''") & "')"

            'Change to SAVE Level 2, Level 1 is always creator


            'ltempsql = "Insert into workflowstatus (wst_workflowid,wst_ucode,wst_refno,wst_module,wst_status,wst_level,wst_merchantid,wst_filtercode,wst_createon,wst_createby,wst_subject) Values " & _
            '"(" & WorkflowID & ",'" & UniqueCode & "','" & RefNo.Replace("'", "''") & "','" & Modulecode & "','Pending',2,'" & WebLib.MerchantID & "','" & WebLib.FilterCode & "',getdate(),'" & WebLib.LoginUser & "','" & DocumentTitle.ToString.Replace("'", "''") & "')"


            'Will only run 1 time, which is the first time
            Dim lnextlevel As String
            Dim ooLevel As New clsworkflow
            lnextlevel = ooLevel.GetNextLevel(WorkflowID, "1")
            ooLevel = Nothing
            If isnumeric(lnextlevel) = False Then
                lnextlevel = 1
            End If

            'Hardcode to 1 after save...need to submit - 20151130
            lnextlevel = 1

            'ltempsql = "Insert into workflowstatus (wst_workflowid,wst_ucode,wst_refno,wst_module,wst_status,wst_level,wst_merchantid,wst_filtercode,wst_createon,wst_createby,wst_subject) Values " & _
            '"(" & WorkflowID & ",'" & UniqueCode & "','" & RefNo.Replace("'", "''") & "','" & Modulecode & "','Pending'," & lnextlevel & ",'" & WebLib.MerchantID & "','" & WebLib.FilterCode & "',getdate(),'" & WebLib.LoginUser & "','" & DocumentTitle.ToString.Replace("'", "''") & "')"

            ltempsql = "Insert into workflowstatus (wst_workflowid,wst_ucode,wst_refno,wst_module,wst_status,wst_level,wst_merchantid,wst_filtercode,wst_createon,wst_createby,wst_subject,wst_param1,wst_param2,wst_param3,wst_param4) Values " & _
            "(" & WorkflowID & ",'" & UniqueCode & "','" & RefNo.Replace("'", "''") & "','" & Modulecode & "','Pending'," & lnextlevel & ",'" & WebLib.MerchantID & "','" & WebLib.FilterCode & "',getdate(),'" & WebLib.LoginUser & "','" & DocumentTitle.ToString.Replace("'", "''") & "','" & utils.checktext(param1) & "','" & utils.checktext(param2) & "','" & utils.checktext(param3) & "','" & utils.checktext(param4) & "')"

            wwid.value = workflowid
            wucode.value = UniqueCode
            ltempsql = ltempsql & ";" & getAuditSQL("CREATE", RefNo, "", UniqueCode)



        Else
            ltempsql = ""

        End If

        Return ltempsql
    End Function
    Private Sub setstatus(ByVal paStatus As String)
        btnwApprove.Enabled = False
        btnwReject.Enabled = False
        btnwCancel.Enabled = False
        btnAttachments.Enabled = True
        btnTracking.Enabled = True
        btnwresend.visible = False

        If custommode = True Then
            btnwApprove.visible = False
            btnwReject.visible = False
            btnwCancel.visible = False

        Else
            btnwApprove.visible = True
            btnwReject.visible = True
            btnwCancel.visible = True

        End If


        Try
            If weblib.LoginisFullAdmin = True Then
                btnwresend.visible = True
            End If
        Catch ex As Exception

        End Try


        If attachmentbycreatoronly = True Then

            If isnumeric(wlevelA.Value) = True Then
                If CLng(wlevelA.Value) > 1 Then
                    btnAttachments.visible = False
                Else
                    btnAttachments.visible = True

                End If
            End If


        Else
            btnAttachments.visible = True


        End If



        If (workflownamespace & "").trim = "" Then

            btnAttachments.Enabled = False
            btnTracking.Enabled = False
            litdetails.Text = "<font color=""red"">WORKFLOW NAMESPACE NOT DEFINED<font>"
            Exit Sub
        End If

        Select Case paStatus.ToLower

            Case "pending"
                litstatus.Text = "<img src=""" & WebLib.ClientURL("graphics/chops/pending300.png") & """ width=""100%"">"

            Case "approved"
                litstatus.Text = "<img src=""" & WebLib.ClientURL("graphics/chops/approve300.png") & """ width=""100%"">"

            Case "rejected"
                litstatus.Text = "<img src=""" & WebLib.ClientURL("graphics/chops/rejected300.png") & """ width=""100%"">"
            Case "cancelled"
                litstatus.Text = "<img src=""" & WebLib.ClientURL("graphics/chops/cancel300.png") & """ width=""100%"">"
            Case Else
                litstatus.Text = "<img src=""" & WebLib.ClientURL("graphics/chops/pending300.png") & """ width=""100%"">"
                btnAttachments.Enabled = False
                btnTracking.Enabled = False

        End Select
        If paStatus.ToLower <> "pending" Then

            Call GetAttachments()
            Call GetAuditLogs()
            btnAttachments.Visible = False

            Exit Sub


        Else

            Call loadActions(wlevel.Value)
            Call GetApprovalPerson()
            Call GetAttachments()
            Call GetAuditLogs()

        End If

    End Sub
    Private Function getActionActionSQL(ByVal pActionValue As String, ByVal pActionStatus As String) As String
        Dim lLevel As String = ""
        Dim lSQL As String = ""
        lLevel = wlevel.Value

        If IsNumeric(lLevel) = False Then
            lLevel = 0
        End If


        '**** Added for Amount Control ********
        '21082016
        Try

            If pActionStatus.tolower = "approved" Then
                If (wlevelamtenabled.value & "").tolower.trim = "true" Then
                    If isnumeric(wlevelamt.value) = False Then
                        wlevelamt.value = 0
                    End If

                    If CDbl(documentamt) > CDbl(wlevelamt.value) Then
                        ' pActionValue = 2 'Force Proceed next level (cant approve this amount)
                        'No need do anything
                    Else
                        If (wlevelamtend.value & "").tolower.trim = "true" Then
                            pActionValue = 1 'Force End workflow cause amount met
                        End If
                    End If
                End If
            End If
        Catch ex As Exception

        End Try

        '**************************************


        Select Case pActionValue.ToLower

            Case "1"  'End Workflow
                'Remove merchant ID, if not cant update status sometimes
                '                lSQL = "Update WorkflowStatus set wst_lastupdateon=getdate(),wst_status='" & pActionStatus & "' where wst_level=" & lLevel & " and wst_ucode='" & wucode.Value & "' and wst_merchantid='" & WebLib.MerchantID & "' and wst_filtercode='" & WebLib.FilterCode & "'"
                lSQL = "Update WorkflowStatus set wst_lastupdateon=getdate(),wst_status='" & pActionStatus & "' where wst_level=" & lLevel & " and wst_ucode='" & wucode.Value & "' and wst_filtercode='" & WebLib.FilterCode & "'"
                If pActionStatus.tolower = "approved" Then
                    isApproved = True
                End If


            Case "2"  'Proceed Next Approval
                Dim lnextlevel As String
                Dim ooLevel As New clsworkflow
                lnextlevel = ooLevel.GetNextLevel(wwid.value, lLevel)
                ooLevel = Nothing
                If isnumeric(lnextlevel) = False Then
                    lSQL = ""
                Else
                    'Remove merchant ID, if not cant update status sometimes
                    'lSQL = "Update WorkflowStatus set wst_lastupdateon=getdate(), wst_level = " & lnextlevel & " where wst_level=" & lLevel & " and wst_ucode='" & wucode.Value & "' and wst_merchantid='" & WebLib.MerchantID & "' and wst_filtercode='" & WebLib.FilterCode & "'"
                    lSQL = "Update WorkflowStatus set wst_lastupdateon=getdate(), wst_level = " & lnextlevel & " where wst_level=" & lLevel & " and wst_ucode='" & wucode.Value & "' and wst_filtercode='" & WebLib.FilterCode & "'"
                End If

            Case "3"  'Back to Creator
                'Remove merchant ID, if not cant update status sometimes
                'lSQL = "Update WorkflowStatus set wst_lastupdateon=getdate(), wst_level = 1,wst_status='Pending' where wst_level=" & lLevel & " and wst_ucode='" & wucode.Value & "' and wst_merchantid='" & WebLib.MerchantID & "' and wst_filtercode='" & WebLib.FilterCode & "'"
                lSQL = "Update WorkflowStatus set wst_lastupdateon=getdate(), wst_level = 1,wst_status='Pending' where wst_level=" & lLevel & " and wst_ucode='" & wucode.Value & "' and wst_filtercode='" & WebLib.FilterCode & "'"

            Case "4"  'Back to Previous Step
                If CLng(lLevel) > 1 Then
                    Dim lprevlevel As String
                    Dim ooLevel As New clsworkflow
                    lprevlevel = ooLevel.GetPreviousLevel(wwid.value, lLevel)
                    ooLevel = Nothing
                    If isnumeric(lprevlevel) = False Then
                        lSQL = ""
                    Else
                        'Remove merchant ID, if not cant update status sometimes
                        'lSQL = "Update WorkflowStatus set wst_lastupdateon=getdate(), wst_level =" & lprevlevel & " where wst_level=" & lLevel & " and wst_ucode='" & wucode.Value & "' and wst_merchantid='" & WebLib.MerchantID & "' and wst_filtercode='" & WebLib.FilterCode & "'"
                        lSQL = "Update WorkflowStatus set wst_lastupdateon=getdate(), wst_level =" & lprevlevel & " where wst_level=" & lLevel & " and wst_ucode='" & wucode.Value & "' and wst_filtercode='" & WebLib.FilterCode & "'"
                    End If

                End If
        End Select
        Return lSQL

    End Function
    Private Sub getDocs()



    End Sub
    Public Function getapprovesql(Optional ByVal pType As String = "") As String

        Dim lsql As String = ""
        Dim lAction As String = ""


        Select Case pType.tolower

            Case "reject"
                lsql = getActionActionSQL(aRType.Value, "Rejected")
                laction = btnwreject.text
                If (laction & "").trim = "" Then
                    laction = "Reject"
                End If
                'lsql = lsql & ";" & getAuditSQL("Reject", cus_wrefno.Value, "", wucode.Value)
                lsql = lsql & ";" & getAuditSQL(laction, cus_wrefno.Value, "", wucode.Value)

            Case "cancel"
                lsql = getActionActionSQL(aCType.Value, "Cancelled")
                laction = btnwcancel.text
                If (laction & "").trim = "" Then
                    laction = "Cancel"
                End If

                'lsql = lsql & ";" & getAuditSQL("Cancel", cus_wrefno.Value, "", wucode.Value)
                lsql = lsql & ";" & getAuditSQL(laction, cus_wrefno.Value, "", wucode.Value)

            Case "approve"
                lsql = getActionActionSQL(aAType.Value, "Approved")
                laction = btnwapprove.text
                If (laction & "").trim = "" Then
                    laction = "Approve"
                End If


                'lsql = lsql & ";" & getAuditSQL("Approve", cus_wrefno.Value, "", wucode.Value)
                lsql = lsql & ";" & getAuditSQL(laction, cus_wrefno.Value, "", wucode.Value)

            Case Else

        End Select


        Return lsql
    End Function
    Private Sub ApproveMod()
        RaiseEvent OnApproveEvent(Me, New EventArgs())

    End Sub
    Public Function notifyusers(ByVal pAction As String) As Boolean

        Select Case pAction.tolower
            Case "approve"
                If isapproved = True Then
                    pAction = "A"
                Else
                    pAction = "U"
                End If
            Case "route"
                pAction = "U"
            Case "reject"
                pAction = "R"
            Case "cancel"
                pAction = "C"
            Case Else
                ErrorMsg = "Action not defined"
                Return False
                Exit Function

        End Select

        If paction.trim = "" Then
            ErrorMsg = "Action not defined"
            Return False
            Exit Function
        End If
        Dim ooemail As New clsWorkflowEmail

        Dim pLevel As String
        pLevel = wlevel.Value

        If IsNumeric(pLevel) = False Then
            pLevel = 1
        End If
        If IsNumeric(wwid.Value) = False Then
            ErrorMsg = "Wrokflow not defined"
            Return False
            Exit Function
        End If


        Dim lversion2 As Boolean = False
        If (wversion.value & "").trim = "2" Then
            lversion2 = True
        End If

        If ooemail.NotifyUsers(wwid.Value, pLevel, pAction, workflownamespace, wucode.value, lversion2) = False Then


            ErrorMsg = ooemail.ErrorMsg
            Return False

            lblMessage.Text = WebLib.getAlertMessageStyle(ooemail.ErrorMsg)
            Exit Function
        End If

        Try


            If pAction = "U" Then  '****** Added on 2016/10/20 to stop sending notification if is not route to top
                If ooemail.NotifyUsers(wwid.Value, pLevel, "N", workflownamespace, wucode.value) = False Then
                    ErrorMsg = ooemail.ErrorMsg
                    Return False

                    lblMessage.Text = WebLib.getAlertMessageStyle(ooemail.ErrorMsg)
                    Exit Function
                End If
            End If  '**** End of Add
        Catch ex As Exception

        End Try

        ooemail = Nothing
        Return True
    End Function

    Public Sub resend(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim dr As DataRow
        Dim lversion2 As Boolean = False
        Dim pLevel As String

        If weblib.LoginIsFullAdmin = False Then
            lblMessage.Text = WebLib.getAlertMessageStyle("Invalid User Rights")
            Exit Sub
        End If

        Dim ooLevel As New clsworkflow
        pLevel = ooLevel.GetPreviousLevel(wwid.value, wlevel.Value)
        ooLevel = Nothing

        If IsNumeric(pLevel) = False Then
            lblMessage.Text = WebLib.getAlertMessageStyle("Invalid Level")
            Exit Sub
        End If
        If IsNumeric(wwid.Value) = False Then
            lblMessage.Text = WebLib.getAlertMessageStyle("Wrokflow not defined")
            Exit Sub
        End If


        Try
            cmd.CommandText = "Select top 1 wui_emailSf from workflowitems where wui_wid=" & wwid.Value & " and isnull(wui_no,0)=" & pLevel & " order by wui_no asc"
            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")
            For Each dr In ds.Tables("datarecords").Rows

                Try
                    If weblib.bittoboolean(dr("wui_emailSf") & "") = True Then
                        lversion2 = True
                    Else
                        lversion2 = False
                    End If
                Catch ex As Exception
                    lblMessage.Text = WebLib.getAlertMessageStyle(ex.Message)
                    Exit Sub
                End Try


                Exit For
            Next
            cn.Close()
            cmd.Dispose()
            cn.Dispose()


        Catch ex As Exception
            lblMessage.Text = WebLib.getAlertMessageStyle(ex.Message)
            Exit Sub
        End Try


        Dim ooemail As New clsWorkflowEmail

  

        If ooemail.NotifyUsers(wwid.Value, pLevel, "U", workflownamespace, wucode.value, lversion2) = False Then
            ErrorMsg = ooemail.ErrorMsg
            lblMessage.Text = WebLib.getAlertMessageStyle(ooemail.ErrorMsg)
        Else
            lblMessage.Text = WebLib.getAlertMessageStyle("Email Resent Successful")
        End If

        ooemail = Nothing

    End Sub
    Public Sub approve(ByVal sender As System.Object, ByVal e As System.EventArgs)

        If overridemode = True Then
            Call approvemod()
            Exit Sub
        End If


        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()
        Dim counter As Integer = 0

        If IsNumeric(wwid.Value) = False Then
            Exit Sub
        End If
        Dim lsql As String
        '        lsql = getActionActionSQL(aAType.Value, "Approved")
        '       lsql = lsql & ";" & getAuditSQL("Approve", cus_wrefno.Value, "", wucode.Value)
        lsql = getapprovesql("approve")
        Try
            cn.Open()
            cmd.CommandText = lsql
            cmd.Connection = cn
            cmd.ExecuteNonQuery()
            cn.Close()
            cmd.Dispose()
            cn.Dispose()
            'Call LoadData()
            '            response.redirect(redirector.
            Call notifyusers("approve")

            response.redirect("~/" & Redirector.Redirect(workflownamespace, wucode.Value, False))


        Catch ex As Exception
            lblMessage.Text = WebLib.getAlertMessageStyle(ex.Message)
        End Try
    End Sub
    Private Sub rejectMod()

        RaiseEvent OnRejectEvent(Me, New EventArgs())

    End Sub
    Public Sub reject(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If overridemode = True Then
            Call rejectmod()
            Exit Sub
        End If

        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()
        Dim counter As Integer = 0

        If IsNumeric(wwid.Value) = False Then
            Exit Sub
        End If
        Dim lsql As String
        '        lsql = getActionActionSQL(aRType.Value, "Rejected")
        '       lsql = lsql & ";" & getAuditSQL("Reject", cus_wrefno.Value, "", wucode.Value)
        lsql = getapprovesql("reject")
        Try
            cn.Open()
            cmd.CommandText = lsql
            cmd.Connection = cn
            cmd.ExecuteNonQuery()
            cn.Close()
            cmd.Dispose()
            cn.Dispose()
            '            Call LoadData()
            Call notifyusers("reject")

            response.redirect("~/" & Redirector.Redirect(workflownamespace, wucode.Value, False))

        Catch ex As Exception
            lblMessage.Text = WebLib.getAlertMessageStyle(ex.Message)
        End Try

    End Sub
    Private Sub CancelMod()
        RaiseEvent OnCancelEvent(Me, New EventArgs())

    End Sub
    Public Sub cancel(ByVal sender As System.Object, ByVal e As System.EventArgs)

        If overridemode = True Then

            Call cancelmod()
            Exit Sub
        End If


        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()
        Dim counter As Integer = 0

        If IsNumeric(wwid.Value) = False Then
            Exit Sub
        End If
        Dim lsql As String
        '        lsql = getActionActionSQL(aCType.Value, "Cancelled")
        '       lsql = lsql & ";" & getAuditSQL("Cancel", cus_wrefno.Value, "", wucode.Value)
        lsql = getapprovesql("cancel")
        Try
            cn.Open()
            cmd.CommandText = lsql
            cmd.Connection = cn
            cmd.ExecuteNonQuery()
            cn.Close()
            cmd.Dispose()
            cn.Dispose()

            Call notifyusers("cancel")

            '            Call LoadData()
            response.redirect("~/" & Redirector.Redirect(workflownamespace, wucode.Value, False))

        Catch ex As Exception
            lblMessage.Text = WebLib.getAlertMessageStyle(ex.Message)
        End Try
    End Sub
    Public Sub viewattach(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '        If wucode.Value.Trim = "" Then
        'lblMessage.Text = " Inconsistency Detected. This action is not allowed"
        'Exit Sub
        'End If
        If IsNumeric(wwid.Value) = False Or wucode.Value.Trim = "" Then
            lblMessage.Text = " Inconsistency Detected. This action is not allowed"
            Exit Sub
        End If


        Response.Redirect("postpage.aspx?NextPage=" & WebLib.ClientURL("modules/docrepo/docrepomod.aspx") & "&ba=" & wucode.Value)
    End Sub
    Public Sub viewworkflow(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If IsNumeric(wwid.Value) = False Or wucode.Value.Trim = "" Then
            lblMessage.Text = " Inconsistency Detected. This action is not allowed"
            Exit Sub
        End If

        Response.Redirect("postpage.aspx?NextPage=" & WebLib.ClientURL("modules/workflow/wbuilderpreview.aspx") & "&ga=" & wwid.Value & " &ba=" & wucode.Value)
    End Sub
    Private Function canAction() As Boolean
        If (wr.Value & "").Trim = "" Then
            Return False
            Exit Function
        End If

        If (wrm.Value & "").Trim = "" Then
            Return False
            Exit Function
        End If
        Dim lrightstring As String = "|" & wrm.Value.Replace(";;", ";;|")

        If Right(lrightstring, 1) = "|" Then
            lrightstring = Left(lrightstring, lrightstring.Length - 1)
        End If

        Dim temp
        Dim lvalue As String = ""
        temp = Microsoft.VisualBasic.Split(wr.Value.Replace(" ", ""), ";;")

        If UBound(temp) = 0 Then
            lvalue = wr.Value.Replace(" ", "")
            If InStr(1, lrightstring, "|" & lvalue & ";;") < 1 Then
                Return False
                Exit Function
            End If
        Else
            Dim counter As Integer
            For counter = 0 To UBound(temp)
                lvalue = temp(counter)
                If InStr(1, lrightstring, "|" & lvalue & ";;") >= 1 Then
                    Return True
                    Exit Function
                End If

            Next

        End If


    End Function
    Private Sub GetRights()

        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim dr As DataRow
        Dim ltemp As String = ""

        Dim luserid As String = ViFeandi.General.GetValue(connectionstring, "secuserinfo", "usr_id", "usr_code", "'" & WebLib.LoginUser & "'", "", "", "", "") & ""

        If IsNumeric(luserid) = False Then
            luserid = 0
        End If

        ltemp = "SELECT CAST((select ltrim(convert(varchar(max),wur_wgroupid) + ';;') as 'data()' from wgrouprights Where wur_uid=" & luserid & " for xml path('')) AS VARCHAR(MAX)) AS RtnData"


        cn.open()
        cmd.CommandText = ltemp
        cmd.Connection = cn
        ad.SelectCommand = cmd
        ad.Fill(ds, "datarecords")
        For Each dr In ds.Tables("datarecords").Rows
            counter = counter + 1
            wr.Value = dr("RtnData") & ""
        Next

        cn.Close()
        cmd.dispose()
        cn.dispose()

    End Sub
    Private Function GetApprovalPerson() As String
        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim dr As DataRow
        Dim lstring As String = ""
        Dim lrights As String = ""
        lrights = (wrm.Value & "").replace(";;", ",")


        If lrights.trim <> "" Then
            lrights = lrights & "0"
        End If

        litapprovalperson.text = ""

        If lrights.trim = "" Then
            Return ""
            Exit Function
        End If

        Try
            cmd.CommandText = "Select distinct usr_name from secuserinfo inner join wgrouprights on secuserinfo.usr_id = wur_uid and wur_wgroupid in (" & lrights & ") "
            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")
            For Each dr In ds.Tables("datarecords").Rows
                lstring = lstring & dr("usr_name") & "<br>"
            Next
            cn.Close()
            cmd.Dispose()
            cn.Dispose()

            If lstring.trim <> "" Then
                lstring = "<font color=""blue"">" & lstring & "</font><br><br>"
            End If

            litapprovalperson.text = "<b>Authorised Approval by : </b><br>" & lstring
            Return lstring
        Catch ex As Exception
            lblMessage.Text = WebLib.getAlertMessageStyle(ex.Message)
            Return ""

        End Try



    End Function
    Private Function GetAuditLogs() As String
        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim dr As DataRow
        Dim lstring As String = ""

        litAudit.text = ""

        Try
            cmd.CommandText = "Select * from workflowaudit where wfa_ucode='" & wucode.Value & "' order by wfa_createon asc"
            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")
            For Each dr In ds.Tables("datarecords").Rows
                lstring = lstring & dr("wfa_createon").ToString.ToUpper & " : <b><br>" & dr("wfa_code").ToString.ToUpper & "</b> by " & dr("wfa_createby") & "<br>"
            Next
            cn.Close()
            cmd.Dispose()
            cn.Dispose()

            If lstring.trim <> "" Then
                lstring = "<font color=""black"">" & lstring & "</font><br><br>"
            End If

            litAudit.text = "<b>Audit Logs</b><br>" & lstring
            Return ""
        Catch ex As Exception
            lblMessage.Text = WebLib.getAlertMessageStyle(ex.Message)
            Return ""

        End Try



    End Function
    Private Sub GetAttachments()
        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim dr As DataRow

        Dim lstring As String = ""
        litattachments.text = ""

        If (wucode.Value & "").trim = "" Then
            Exit Sub
        End If


        Try

            Dim obj As New clsFileTypes

            cmd.CommandText = "Select docdoc.* from docdoc left outer join docgroup on doc_group = dg_id where isnull(doc_group,0) = -1 and doc_uniqueid='" & wucode.Value & "' order by doc_subject"
            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")
            For Each dr In ds.Tables("datarecords").Rows
                obj.InitFile(dr("doc_attach1") & "")
                lstring = lstring & "<table width=""100%""><tr><td width=""20%"" valign=""top""><img src=""" & WebLib.ClientURL(obj.FileImageFile) & """ width=""100%""></td><td width=""80%"" valign=""top"" class=""cssdetail""><b><a href=""#"" onClick=""$.colorbox({iframe:true,opacity:0.5,trapFocus:true,href: '" & WebLib.ClientURL(dr("doc_attach1path").ToString.Trim & dr("doc_attach1").ToString.Trim) & "',width:'90%',height:'90%'})"">" & dr("doc_subject") & "</a></b><br><font color=""gray"">" & obj.FileType & "</font></td></tr></table>"
            Next
            cn.Close()
            cmd.Dispose()
            cn.Dispose()
            obj = Nothing
            If lstring.trim <> "" Then
                litattachments.text = lstring & "<br><br>"

            Else
                litattachments.text = "<font color=""red"">No Attachments</font>"

            End If
            '            Return lstring
        Catch ex As Exception
            lblMessage.Text = WebLib.getAlertMessageStyle(ex.Message)

        End Try



    End Sub
    Private Sub loadActions(ByVal pLevel As String)
        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim dr As DataRow
        Dim bcan As Boolean = True
        If IsNumeric(pLevel) = False Then
            Exit Sub

        End If
        If IsNumeric(wwid.Value) = False Then
            Exit Sub
        End If

        Try
            cmd.CommandText = "Select top 1 * from workflowitems where wui_wid=" & wwid.Value & " and isnull(wui_no,0)=" & pLevel & " order by wui_no asc"
            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")
            For Each dr In ds.Tables("datarecords").Rows
                If WebLib.BitToBoolean(dr("wui_approve") & "") = True Then
                    btnwApprove.Enabled = True
                End If
                If WebLib.BitToBoolean(dr("wui_cancel") & "") = True Then
                    btnwCancel.Enabled = True
                End If
                If WebLib.BitToBoolean(dr("wui_reject") & "") = True Then
                    btnwReject.Enabled = True
                End If

                If WebLib.BitToBoolean(dr("wui_allowattach") & "") = True Then
                    btnattachments.Visible = True
                    btnattachments.Enabled = True
                End If


                'Added for customise approve button
                If (dr("wui_approvename") & "").trim <> "" Then
                    btnwApprove.Text = (dr("wui_approvename") & "").trim
                End If

                If (dr("wui_cancelname") & "").trim <> "" Then
                    btnwCancel.Text = (dr("wui_cancelname") & "").trim
                End If

                If (dr("wui_rejectname") & "").trim <> "" Then
                    btnwReject.Text = (dr("wui_rejectname") & "").trim
                End If


                wrm.Value = dr("wui_rights") & ""


                aCType.Value = dr("wui_cancelstep") & ""
                aRType.Value = dr("wui_rejectstep") & ""
                aAType.Value = dr("wui_approvestep") & ""



                Try
                    wlevelamtenabled.value = WebLib.Bittoboolean(dr("wui_approveval") & "")
                    wlevelamt.value = dr("wui_approvevalamt") & ""
                    wlevelamtend.value = WebLib.Bittoboolean(dr("wui_approvalvalend") & "")
                Catch ex As Exception

                End Try


                Try
                    If weblib.bittoboolean(dr("wui_emailSf") & "") = True Then
                        wversion.value = 2
                    Else
                        wversion.value = ""
                    End If
                Catch ex As Exception

                End Try


                Exit For
            Next
            cn.Close()
            cmd.Dispose()
            cn.Dispose()


            If canAction() = False Then
                btnwApprove.Enabled = False
                btnwCancel.Enabled = False
                btnwReject.Enabled = False
            End If


        Catch ex As Exception
            lblMessage.Text = WebLib.getAlertMessageStyle(ex.Message)
        End Try
    End Sub
    Private Function LoadAreaButton()

        Dim obj As Object
        'obj = Page.FindControl("phlevel" & wlevel.Value)
        obj = Page.FindControl("phlevel" & wlevelA.Value)

        If Not obj Is Nothing Then

            Dim objlit As New Literal
            'objlit.ID = "litID" & wlevel.Value
            objlit.ID = "litID" & wlevelA.Value

            objlit.text = "&nbsp;&nbsp;"
            obj.Controls.Add(objlit)


            If btnwApprove.Enabled = True Then
                Dim obj1 As New button
                'obj1.ID = "dobj" & wlevel.Value & "A"
                obj1.ID = "dobj" & wlevelA.Value & "A"

                obj1.Style.Add("width", "100")
                obj1.Text = btnwApprove.text
                obj1.Style.Add("height", "25")
                AddHandler obj1.Click, AddressOf Me.approvemod
                obj.Controls.Add(obj1)


                objlit = New Literal
                'objlit.ID = "litID" & wlevel.Value & "1"
                objlit.ID = "litID" & wlevelA.Value & "1"

                objlit.text = "&nbsp;"
                obj.Controls.Add(objlit)


            End If

            If btnwReject.Enabled = True Then
                Dim obj1 As New button
                'obj1.ID = "dobj" & wlevel.Value & "R"
                obj1.ID = "dobj" & wlevelA.Value & "R"

                obj1.Style.Add("width", "100")
                obj1.Text = btnwReject.text
                obj1.Style.Add("height", "25")
                AddHandler obj1.Click, AddressOf Me.rejectmod
                obj.Controls.Add(obj1)

                objlit = New Literal
                'objlit.ID = "litID" & wlevel.Value & "2"
                objlit.ID = "litID" & wlevelA.Value & "2"

                objlit.text = "&nbsp;"
                obj.Controls.Add(objlit)


            End If
            If btnwCancel.Enabled = True Then
                Dim obj1 As New button
                'obj1.ID = "dobj" & wlevel.Value & "C"
                objlit.ID = "litID" & wlevelA.Value & "2"

                obj1.Style.Add("width", "100")
                obj1.Text = btnwCancel.text
                obj1.Style.Add("height", "25")
                AddHandler obj1.Click, AddressOf Me.cancelmod
                obj.Controls.Add(obj1)
                obj.Controls.Add(objlit)

            End If

        End If

    End Function
    Private Function LoadAudit() As Boolean
        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim dr As DataRow
        Dim templi As ListItem
        Dim obj As Object
        Try
            cmd.CommandText = "Select wfa_code,wfa_createon,wfa_createby,wfa_level from workflowaudit where wfa_ucode='" & wucode.Value & "' order by wfa_id asc"
            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")
            For Each dr In ds.Tables("datarecords").Rows
                obj = Page.FindControl("lbllevel" & (dr("wfa_level") & "").trim)
                If Not obj Is Nothing Then
                    obj.text = "&nbsp;&nbsp;<font color=""blue"" class=""cssdetail"">" & (dr("wfa_code") & "").tostring.toupper & " by " & (dr("wfa_createby") & "").tostring.toupper & " " & (dr("wfa_createon") & "").tostring.toupper & "</font>"
                End If
            Next
            cn.Close()
            cmd.Dispose()
            cn.Dispose()
        Catch ex As Exception
            lblMessage.Text = WebLib.getAlertMessageStyle(ex.Message)
        End Try

    End Function
    Public Function LoadDataApproval2() As Boolean
        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim dr As DataRow
        Dim templi As ListItem

        Try
            cmd.CommandText = "Select top 1 * from workflowtrack where ws_ucode='" & wucode.Value & "' and isnull(ws_ok,0) = 0 order by ws_wno asc"
            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")
            For Each dr In ds.Tables("datarecords").Rows


                Exit For
            Next
            cn.Close()
            cmd.Dispose()
            cn.Dispose()
        Catch ex As Exception
            lblMessage.Text = WebLib.getAlertMessageStyle(ex.Message)
        End Try
    End Function
    Public Sub EnableDisable(ByRef pParent As Object)

        '**** Added ****
        If StrictValidationGroup = True Then
            Call EnableDisable2(pParent)
            Exit Sub
        End If
        '****


        Dim obj As Object
        Dim llevel As String = ""
        Dim bcan As Boolean = True
        If IsNumeric(wlevelA.Value) = False Then
            'llevel = "0"
            llevel = "1"

        Else
            llevel = wlevelA.Value

            If canAction() = True Then
                bcan = True
            Else
                bcan = False
            End If


        End If



        For Each obj In pParent.Controls

            If TypeOf obj Is TextBox Then
                If obj.validationgroup.ToString.Trim <> "" Then
                    '  If obj.validationgroup.ToString.Trim = llevel & "-" And WorkflowEnded = False And bcan = True Then
                    If InStr(1, obj.validationgroup.ToString.Trim, llevel & "-") >= 1 And WorkflowEnded = False And bcan = True Then

                    Else
                        obj.style.add("background-color", "#EEEEEE")
                        obj.enabled = False
                    End If
                End If
            End If


            If (obj.gettype()).tostring.tolower = "asp.usercontrols_datepicker_ascx" Then
                If obj.validationgroup.ToString.Trim <> "" Then
                    If InStr(1, obj.validationgroup.ToString.Trim, llevel & "-") >= 1 And WorkflowEnded = False And bcan = True Then
                        'If obj.validationgroup.ToString.Trim = llevel & "-" And WorkflowEnded = False And bcan = True Then

                    Else
                        Try

                            obj.TextBox1.style.add("background-color", "#EEEEEE")
                        Catch ex As Exception

                        End Try

                        obj.enabled = False
                    End If
                End If
            End If

            If TypeOf obj Is CheckBox Then
                If obj.validationgroup.ToString.Trim <> "" Then
                    '  If obj.validationgroup.ToString.Trim = llevel & "-" And WorkflowEnded = False And bcan = True Then
                    If InStr(1, obj.validationgroup.ToString.Trim, llevel & "-") >= 1 And WorkflowEnded = False And bcan = True Then

                    Else
                        obj.enabled = False
                    End If
                End If
            End If


            If TypeOf obj Is RadioButtonList Then
                If obj.validationgroup.ToString.Trim <> "" Then
                    'If obj.validationgroup.ToString.Trim = llevel & "-" And WorkflowEnded = False And bcan = True Then
                    If InStr(1, obj.validationgroup.ToString.Trim, llevel & "-") >= 1 And WorkflowEnded = False And bcan = True Then

                    Else
                        obj.enabled = False
                    End If
                End If
            End If


            If TypeOf obj Is DropdownList Then
                If obj.validationgroup.ToString.Trim <> "" Then
                    ' If obj.validationgroup.ToString.Trim = llevel & "-" And WorkflowEnded = False And bcan = True Then
                    If InStr(1, obj.validationgroup.ToString.Trim, llevel & "-") >= 1 And WorkflowEnded = False And bcan = True Then

                    Else
                        obj.enabled = False
                    End If
                End If
            End If


        Next
        Exit Sub

    End Sub

    Protected Sub Page_LoadComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Call LoadAreaButton
    End Sub


    Public Sub EnableDisable2(ByRef pParent As Object)

        '**** Added ****
        If StrictValidationGroup = False Then
            Call EnableDisable(pParent)
            Exit Sub
        End If
        '****


        Dim obj As Object
        Dim llevel As String = ""
        Dim bcan As Boolean = True
        If IsNumeric(wlevelA.Value) = False Then
            llevel = "1"

        Else
            llevel = wlevelA.Value

            If canAction() = True Then
                bcan = True
            Else
                bcan = False
            End If


        End If



        For Each obj In pParent.Controls

            If TypeOf obj Is TextBox Then
                If obj.validationgroup.ToString.Trim <> "" Then
                    If obj.validationgroup.ToString.Trim = llevel & "-" And WorkflowEnded = False And bcan = True And obj.validationgroup.ToString.Length = (llevel & "-").trim.length Then
                        'If InStr(1, obj.validationgroup.ToString.Trim, llevel & "-") >= 1 And WorkflowEnded = False And bcan = True Then

                    Else
                        obj.style.add("background-color", "#EEEEEE")
                        obj.enabled = False
                    End If
                End If
            End If


            If (obj.gettype()).tostring.tolower = "asp.usercontrols_datepicker_ascx" Then
                If obj.validationgroup.ToString.Trim <> "" Then
                    ' If InStr(1, obj.validationgroup.ToString.Trim, llevel & "-") >= 1 And WorkflowEnded = False And bcan = True Then
                    If obj.validationgroup.ToString.Trim = llevel & "-" And WorkflowEnded = False And bcan = True And obj.validationgroup.ToString.Length = (llevel & "-").trim.length Then

                    Else
                        Try

                            obj.TextBox1.style.add("background-color", "#EEEEEE")
                        Catch ex As Exception

                        End Try

                        obj.enabled = False
                    End If
                End If
            End If

            If TypeOf obj Is CheckBox Then
                If obj.validationgroup.ToString.Trim <> "" Then
                    If obj.validationgroup.ToString.Trim = llevel & "-" And WorkflowEnded = False And bcan = True And obj.validationgroup.ToString.Length = (llevel & "-").trim.length Then
                        ' If InStr(1, obj.validationgroup.ToString.Trim, llevel & "-") >= 1 And WorkflowEnded = False And bcan = True Then

                    Else
                        obj.enabled = False
                    End If
                End If
            End If


            If TypeOf obj Is RadioButtonList Then
                If obj.validationgroup.ToString.Trim <> "" Then
                    If obj.validationgroup.ToString.Trim = llevel & "-" And WorkflowEnded = False And bcan = True And obj.validationgroup.ToString.Length = (llevel & "-").trim.length Then
                        ' If InStr(1, obj.validationgroup.ToString.Trim, llevel & "-") >= 1 And WorkflowEnded = False And bcan = True Then

                    Else
                        obj.enabled = False
                    End If
                End If
            End If


            If TypeOf obj Is DropdownList Then
                If obj.validationgroup.ToString.Trim <> "" Then
                    If obj.validationgroup.ToString.Trim = llevel & "-" And WorkflowEnded = False And bcan = True And obj.validationgroup.ToString.Length = (llevel & "-").trim.length Then
                        '  If InStr(1, obj.validationgroup.ToString.Trim, llevel & "-") >= 1 And WorkflowEnded = False And bcan = True Then

                    Else
                        obj.enabled = False
                    End If
                End If
            End If


        Next
        Exit Sub

    End Sub

End Class

