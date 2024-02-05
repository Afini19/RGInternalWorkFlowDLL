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

Public Class Redirector
    Public Shared Property PrevURL1()
        Get
            If (HttpContext.Current.Session("PrevURL1") & "" <> "") Then

                Return HttpContext.Current.Session("PrevURL1")
            Else
                Return ""
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("PrevURL1") = value

        End Set
    End Property
    Public Shared Property PrevURL2()
        Get
            If (HttpContext.Current.Session("PrevURL2") & "" <> "") Then

                Return HttpContext.Current.Session("PrevURL2")
            Else
                Return ""
            End If
        End Get
        Set(ByVal value)
            HttpContext.Current.Session("PrevURL2") = value

        End Set
    End Property

    Public Shared Function Redirect(ByVal RedirectNameSpace As String, ByVal RedirectUID As String, Optional ByVal isAdd As Boolean = False, Optional ByVal RefNo As String = "") As String
        Dim connectionstring As String = System.Configuration.ConfigurationSettings.AppSettings("ConnStr")

        Dim cn As New OleDbConnection(connectionstring)
        Dim cmd As New OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim dr As DataRow
        Dim ltemp As String = ""
        Dim lid As String = ""


        Select Case RedirectNameSpace.tolower
            Case "zcustom_tempcl"
                cmd.CommandText = "Select cus_id as uid from zcustom_tempcl where cus_ucode='" & RedirectUID & "'" ' and cus_merchantid='" & weblib.merchantid & "' and cus_filter='" & weblib.filtercode & "'"
                ltemp = "postpage.aspx?NextPage=modules/custom/tempcl.aspx&ga=<<ID>>"

                If isadd = True Then
                    ltemp = "wiz-postpage.aspx?NextPage=wiz-selectcustomer.aspx&wp1=modules/custom/tempcl.aspx&wp5=" & RefNo
                End If

            Case "zcustom_test"
                cmd.CommandText = "Select cus_id as uid from zcustom_test where cus_ucode='" & RedirectUID & "'" ' and cus_merchantid='" & weblib.merchantid & "' and cus_filter='" & weblib.filtercode & "'"
                ltemp = "postpage.aspx?NextPage=modules/custom/testform.aspx&ga=<<ID>>"

            Case "zcustom_inactive"
                cmd.CommandText = "Select cus_id as uid from zcustom_inactive where cus_ucode='" & RedirectUID & "'" ' and cus_merchantid='" & weblib.merchantid & "' and cus_filter='" & weblib.filtercode & "'"
                ltemp = "postpage.aspx?NextPage=modules/custom/inactive.aspx&ga=<<ID>>"

                If isadd = True Then
                    ltemp = "wiz-postpage.aspx?NextPage=wiz-selectcustomer.aspx&wp1=modules/custom/inactive.aspx&wp5=" & RefNo
                End If



            Case "zcustom_rebate"
                cmd.CommandText = "Select cus_id as uid from zcustom_rebate where cus_ucode='" & RedirectUID & "'" ' and cus_merchantid='" & weblib.merchantid & "' and cus_filter='" & weblib.filtercode & "'"
                ltemp = "postpage.aspx?NextPage=modules/custom/rebate.aspx&ga=<<ID>>"

            Case "zcustom_generic"
                cmd.CommandText = "Select cus_id as uid from zcustom_generic where cus_ucode='" & RedirectUID & "'" ' and cus_merchantid='" & weblib.merchantid & "' and cus_filter='" & weblib.filtercode & "'"
                ltemp = "postpage.aspx?NextPage=modules/custom/cusgeneric.aspx&ga=<<ID>>"

            Case "zcustom_samples"
                '                cmd.CommandText = "Select cus_id as uid from zcustom_samples where cus_ucode='" & RedirectUID & "'" ' and cus_merchantid='" & weblib.merchantid & "' and cus_filter='" & weblib.filtercode & "'"
                '               ltemp = "postpage.aspx?NextPage=modules/custom/samples.aspx&ga=<<ID>>"
                cmd.CommandText = "Select cus_id as uid from zcustom_samples where cus_ucode='" & RedirectUID & "'" ' and cus_merchantid='" & weblib.merchantid & "' and cus_filter='" & weblib.filtercode & "'"
                ltemp = "postpage.aspx?NextPage=modules/custom/samples.aspx&ga=<<ID>>"

                If isadd = True Then
                    ltemp = "wiz-postpage.aspx?NextPage=wiz-selectcustomer.aspx&wp1=modules/custom/samples.aspx&wp5=" & RefNo
                End If


            Case "zcustom_ccr"
                cmd.CommandText = "Select cus_id as uid from zcustom_ccr where cus_ucode='" & RedirectUID & "'" ' and cus_merchantid='" & weblib.merchantid & "' and cus_filter='" & weblib.filtercode & "'"
                ltemp = "postpage.aspx?NextPage=modules/custom/cusccrST.aspx&events=" & RefNo & "&ga=<<ID>>"

                If isadd = True Then
                    ltemp = "wiz-postpage.aspx?NextPage=wiz-selectcustomer.aspx&wp1=modules/custom/cusccrST.aspx&wp7=Y&wp5=" & RefNo
                End If

            Case "zcustom_ccrp"
                cmd.CommandText = "Select cus_id as uid from zcustom_ccr where cus_ucode='" & RedirectUID & "'" ' and cus_merchantid='" & weblib.merchantid & "' and cus_filter='" & weblib.filtercode & "'"
                ltemp = "postpage.aspx?NextPage=modules/custom/cusccrST.aspx&events=" & RefNo & "&ga=<<ID>>"

                If isadd = True Then
                    ltemp = "wiz-postpage.aspx?NextPage=wiz-selectcustomer.aspx&wp1=modules/custom/cusccrST.aspx&wp7=Y&wp5=" & RefNo
                End If

            Case "zcustom_ccrs"
                cmd.CommandText = "Select cus_id as uid from zcustom_ccr where cus_ucode='" & RedirectUID & "'" ' and cus_merchantid='" & weblib.merchantid & "' and cus_filter='" & weblib.filtercode & "'"
                ltemp = "postpage.aspx?NextPage=modules/custom/cusccrST.aspx&events=" & RefNo & "&ga=<<ID>>"

                If isadd = True Then
                    ltemp = "wiz-postpage.aspx?NextPage=wiz-selectcustomer.aspx&wp1=modules/custom/cusccrST.aspx&wp7=Y&wp5=" & RefNo
                End If

                '            Case "zcustom_climit"
                '               cmd.CommandText = "Select cus_id as uid from zcustom_climit where cus_ucode='" & RedirectUID & "'" ' and cus_merchantid='" & weblib.merchantid & "' and cus_filter='" & weblib.filtercode & "'"
                '              ltemp = "postpage.aspx?NextPage=modules/custom/climit.aspx&ga=<<ID>>"

            Case "zcustom_dncn"
                cmd.CommandText = "Select cus_id as uid from zcustom_dncn where cus_ucode='" & RedirectUID & "'" ' and cus_merchantid='" & weblib.merchantid & "' and cus_filter='" & weblib.filtercode & "'"
                ltemp = "postpage.aspx?NextPage=modules/custom/dncn.aspx&ga=<<ID>>"

                If isadd = True Then
                    ltemp = "wiz-postpage.aspx?NextPage=wiz-selectcustomer.aspx&wp1=modules/custom/dncn.aspx&wp5=" & RefNo
                End If

            Case "zcustom_cn"
                cmd.CommandText = "Select cus_id as uid from zcustom_dncn where cus_ucode='" & RedirectUID & "' and cus_type='CN'" ' and cus_merchantid='" & weblib.merchantid & "' and cus_filter='" & weblib.filtercode & "'"
                ltemp = "postpage.aspx?NextPage=modules/custom/cn.aspx&ga=<<ID>>"

                If isadd = True Then
                    ltemp = "wiz-postpage.aspx?NextPage=wiz-selectcustomer.aspx&wp1=modules/custom/cn.aspx&wp5=" & RefNo
                End If

            Case "zcustom_cn2"
                cmd.CommandText = "Select cus_id as uid from zcustom_dncn where cus_ucode='" & RedirectUID & "' and cus_type='CN2'" ' and cus_merchantid='" & weblib.merchantid & "' and cus_filter='" & weblib.filtercode & "'"
                ltemp = "postpage.aspx?NextPage=modules/custom/cn2.aspx&ga=<<ID>>"

                If isadd = True Then
                    ltemp = "wiz-postpage.aspx?NextPage=wiz-selectcustomer.aspx&wp1=modules/custom/cn2.aspx&wp5=" & RefNo
                End If

            Case "zcustom_cn3"
                cmd.CommandText = "Select cus_id as uid from zcustom_dncn where cus_ucode='" & RedirectUID & "' and cus_type='CN3'" ' and cus_merchantid='" & weblib.merchantid & "' and cus_filter='" & weblib.filtercode & "'"
                ltemp = "postpage.aspx?NextPage=modules/custom/cn3.aspx&ga=<<ID>>"

                If isadd = True Then
                    ltemp = "wiz-postpage.aspx?NextPage=wiz-selectcustomer.aspx&wp1=modules/custom/cn3.aspx&wp5=" & RefNo
                End If



            Case "zcustom_dn"
                cmd.CommandText = "Select cus_id as uid from zcustom_dncn where cus_ucode='" & RedirectUID & "' and cus_type='DN'" ' and cus_merchantid='" & weblib.merchantid & "' and cus_filter='" & weblib.filtercode & "'"
                ltemp = "postpage.aspx?NextPage=modules/custom/dn.aspx&ga=<<ID>>"

                If isadd = True Then
                    ltemp = "wiz-postpage.aspx?NextPage=wiz-selectcustomer.aspx&wp1=modules/custom/dn.aspx&wp5=" & RefNo
                End If


            Case "zcustom_dn2"
                cmd.CommandText = "Select cus_id as uid from zcustom_dncn where cus_ucode='" & RedirectUID & "' and cus_type='DN2'" ' and cus_merchantid='" & weblib.merchantid & "' and cus_filter='" & weblib.filtercode & "'"
                ltemp = "postpage.aspx?NextPage=modules/custom/dn2.aspx&ga=<<ID>>"

                If isadd = True Then
                    ltemp = "wiz-postpage.aspx?NextPage=wiz-selectcustomer.aspx&wp1=modules/custom/dn2.aspx&wp5=" & RefNo
                End If

            Case "zcustom_dn3"
                cmd.CommandText = "Select cus_id as uid from zcustom_dncn where cus_ucode='" & RedirectUID & "' and cus_type='DN3'" ' and cus_merchantid='" & weblib.merchantid & "' and cus_filter='" & weblib.filtercode & "'"
                ltemp = "postpage.aspx?NextPage=modules/custom/dn3.aspx&ga=<<ID>>"

                If isadd = True Then
                    ltemp = "wiz-postpage.aspx?NextPage=wiz-selectcustomer.aspx&wp1=modules/custom/dn3.aspx&wp5=" & RefNo
                End If

            Case "zcustom_clexceed"
                cmd.CommandText = "Select cus_id as uid from zcustom_climit where cus_ucode='" & RedirectUID & "'" ' and cus_merchantid='" & weblib.merchantid & "' and cus_filter='" & weblib.filtercode & "'"
                ltemp = "postpage.aspx?NextPage=modules/custom/climit.aspx&ga=<<ID>>"

                If isadd = True Then
                    ltemp = "wiz-postpage.aspx?NextPage=wiz-selectcustomer.aspx&wp1=modules/custom/climit.aspx&wp5=" & RefNo
                End If


            Case "zcustom_unblockacct"
                cmd.CommandText = "Select cus_id as uid from zcustom_unblockacct where cus_ucode='" & RedirectUID & "'" ' and cus_merchantid='" & weblib.merchantid & "' and cus_filter='" & weblib.filtercode & "'"
                ltemp = "postpage.aspx?NextPage=modules/custom/unblockacct.aspx&ga=<<ID>>"

                If isadd = True Then
                    ltemp = "wiz-postpage.aspx?NextPage=wiz-selectcustomer.aspx&wp1=modules/custom/unblockacct.aspx&wp5=" & RefNo
                End If

            Case "zcustom_ceval"
                cmd.CommandText = "Select cus_id as uid from zcustom_ceval where cus_ucode='" & RedirectUID & "'" ' and cus_merchantid='" & weblib.merchantid & "' and cus_filter='" & weblib.filtercode & "'"
                ltemp = "postpage.aspx?NextPage=modules/custom/ceval.aspx&ga=<<ID>>"

                If isadd = True Then
                    ltemp = "wiz-postpage.aspx?NextPage=wiz-selectcustomer.aspx&wp1=modules/custom/ceval.aspx&wp7=Y&wp5=" & RefNo
                End If

            Case "zcustom_ccrc"
                cmd.CommandText = "Select cus_id as uid from zcustom_ccr where cus_ucode='" & RedirectUID & "'"
                ltemp = "postpage.aspx?NextPage=modules/custom/cusccrc.aspx&ga=<<ID>>"

                If isadd = True Then
                    ltemp = "wiz-postpage.aspx?NextPage=wiz-selectcustomer.aspx&wp1=modules/custom/cusccrc.aspx&wp7=Y&wp5=" & RefNo
                End If

            Case "zcustom_crc"
                cmd.CommandText = "Select cus_id as uid from zcustom_crC where cus_ucode='" & RedirectUID & "'"
                ltemp = "postpage.aspx?NextPage=modules/custom/crC.aspx&ga=<<ID>>"

                If isadd = True Then
                    'ltemp = "wiz-postpage.aspx?NextPage=wiz-selectcustomer.aspx&wp1=modules/custom/cr.aspx&wp7=Y&wp5=" & RefNo
                    ltemp = "wiz-postpage.aspx?NextPage=modules/custom/crC.aspx"
                End If

            Case "zcustom_cri"
                cmd.CommandText = "Select cus_id as uid from zcustom_crI where cus_ucode='" & RedirectUID & "'"
                ltemp = "postpage.aspx?NextPage=modules/custom/crI.aspx&ga=<<ID>>"

                If isAdd = True Then
                    'ltemp = "wiz-postpage.aspx?NextPage=wiz-selectcustomer.aspx&wp1=modules/custom/cr.aspx&wp7=Y&wp5=" & RefNo
                    ltemp = "wiz-postpage.aspx?NextPage=modules/custom/crI.aspx"
                End If

            Case "zcustom_crs"
                cmd.CommandText = "Select cus_id as uid from zcustom_crS where cus_ucode='" & RedirectUID & "'"
                ltemp = "postpage.aspx?NextPage=modules/custom/crS.aspx&ga=<<ID>>"

                If isAdd = True Then
                    'ltemp = "wiz-postpage.aspx?NextPage=wiz-selectcustomer.aspx&wp1=modules/custom/cr.aspx&wp7=Y&wp5=" & RefNo
                    ltemp = "wiz-postpage.aspx?NextPage=modules/custom/crS.aspx"
                End If

            Case "zcustom_sti"
                cmd.CommandText = "Select cus_id as uid from zcustom_stI where cus_ucode='" & RedirectUID & "'"
                ltemp = "postpage.aspx?NextPage=modules/custom/stI.aspx&ga=<<ID>>"

                If isAdd = True Then
                    'ltemp = "wiz-postpage.aspx?NextPage=wiz-selectcustomer.aspx&wp1=modules/custom/cr.aspx&wp7=Y&wp5=" & RefNo
                    ltemp = "wiz-postpage.aspx?NextPage=modules/custom/stI.aspx"
                End If

            Case "zcustom_sts"
                cmd.CommandText = "Select cus_id as uid from zcustom_stS where cus_ucode='" & RedirectUID & "'"
                ltemp = "postpage.aspx?NextPage=modules/custom/stS.aspx&ga=<<ID>>"

                If isAdd = True Then
                    'ltemp = "wiz-postpage.aspx?NextPage=wiz-selectcustomer.aspx&wp1=modules/custom/cr.aspx&wp7=Y&wp5=" & RefNo
                    ltemp = "wiz-postpage.aspx?NextPage=modules/custom/stS.aspx"
                End If

            Case Else
                Return ""
        End Select


        If isadd = False Then

            cn.open()

            cmd.Connection = cn
            ad.SelectCommand = cmd
            ad.Fill(ds, "datarecords")
            For Each dr In ds.Tables("datarecords").Rows
                counter = counter + 1
                lid = dr("uid") & ""
                Exit For
            Next
            cn.Close()

            Return ltemp.replace("<<ID>>", lid)
        Else
            Return ltemp.replace("<<ID>>", "")

        End If


    End Function
End Class