Imports System
Imports System.IO
Imports System.Data
Imports System.Data.ODBC
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

Public Class RuntimeProduct
    Public PartName As String
    Public PartUOM As String
    Public PartNo As String
    Public PartListPrice As String

    Public Function getProductInfo(ByVal _p_productcode As String, Optional ByVal _p_shiptostate As String = "") As Boolean


        Dim conn As OdbcConnection
        Dim comm As OdbcCommand
        Dim dr As DataRow
        Dim counter As Integer = 0
        Dim ad As New Odbc.OdbcDataAdapter
        Dim ds As New DataSet()
        Dim connectionString As String
        Dim sql As String

        Try


            connectionString = Weblib.connEpicor

            If _p_shiptostate = "" Then
                sql = "Select SalesUM,IUM,PartNum,PartDescription from part where partNum='" & _p_productcode.replace("'", "''") & "'"
            Else
                sql = "Select SalesUM,IUM,part.PartNum,PartDescription,PriceLstParts.BasePrice from part inner join Erp.PriceLstParts PriceLstParts on part.partnum = PriceLstParts.partnum inner join PriceLst PriceLst on PriceLstParts.ListCode = PriceLst.ListCode where part.partNum='" & _p_productcode.replace("'", "''") & "' and PriceLst.Character01='" & _p_shiptostate.replace("'", "''") & "' and '" & WebLib.formatthedate(DateTime.Today(),true) & "' between PriceLst.StartDate and ISNULL(PriceLst.EndDate,'" & WebLib.formatthedate(DateTime.Today(),true) & "') "
            End If

            conn = New OdbcConnection(connectionString)
            conn.Open()
            comm = New OdbcCommand(sql, conn)

            ad.SelectCommand = comm
            ad.Fill(ds, "datarecords")
            For Each dr In ds.Tables("datarecords").Rows
                PartName = dr("partdescription") & ""
                PartUOM = dr("SalesUM") & ""
                PartNo = dr("PartNum") & ""
                
                If _p_shiptostate <> "" Then
                    PartListPrice = dr("BasePrice") & ""
                End If

                counter = counter + 1

            Next
            conn.Close()
            comm.Dispose()
            conn.Dispose()

            If counter = 0 Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

End Class

Public Class RuntimeOrderItems

    Public Function getOrderItems(ByVal _p_UID As String) As dataset
        Dim connectionstring As String = System.Configuration.ConfigurationSettings.AppSettings("ConnStr")
        Dim cn As New OleDbConnection(connectionstring)

        Dim cmd As New OleDbCommand()
        Dim ad As New OleDb.OleDbDataAdapter()
        Dim ds As New DataSet()
        Dim counter As Integer = 0
        Dim dr As DataRow


        cn.Open()
        cmd.CommandText = "Select * from salesdetails where sd_uid='" & _p_uid & "'"
        cmd.Connection = cn
        ad.SelectCommand = cmd
        ad.Fill(ds, "datarecords")

        Return ds

    End Function

End Class

Public Class RuntimeCustomer
    Public CustName As String
    Public CustID As String
    Public CustNum As String
    Public Address1 As String
    Public Address2 As String
    Public Address3 As String
    Public City As String
    Public State As String
    Public PostCode As String
    Public hasbranch As Boolean
    Public TermsDescription As String
    Public TermsCode As String
    Public TCLAmount As String
    Public InvoiceCredit As String
    Public InvoiceCreditUnposted As String
    Public ShipAmt As String
    Public OrderCredit As String
    Public CreditLimit As String
    Public OpenInvoice As String
    Public OpenOrders As String
    Public TotalCredit As String
    Public TotalCredit2 As String
    Public CreditPercent2 As String   'For Staff View
    Public CreditPercent As String ' For Customer View
    Public Function getInfo(ByVal _p_parameter As String, Optional ByVal _p_parameter2 As String = "") As Boolean


        Dim conn As OdbcConnection
        Dim comm As OdbcCommand
        Dim dr As DataRow
        Dim counter As Integer = 0
        Dim ad As New Odbc.OdbcDataAdapter
        Dim ds As New DataSet()
        Dim connectionString As String
        Dim sql As String

        Try

            'select CustID,CustNum,Name,Address1,Address2,Address3,City,State,Zip,Customer.checkbox01,Terms.Description  from Customer left outer join Terms  on Customer.Company = Terms.Company and Terms.TermsCode = customer.TermsCode 

            connectionString = Weblib.connEpicor
            'sql = "Select CustID,CustNum,Name,Address1,Address2,Address3,City,State,Zip,checkbox01 from Customer where CustID='" & _p_parameter.replace("'", "''") & "'"
            'sql = "select CustID,CustNum,Name,Address1,Address2,Address3,City,State,Zip,Customer.checkbox01,Terms.Description as TermsDesc,CreditLimit,customer.TermsCode,Customer.Number02 as TCLAmt from Customer left outer join Terms on Customer.Company = Terms.Company and Terms.TermsCode = customer.TermsCode where CustID='" & _p_parameter.replace("'", "''") & "'"
            ' E10
            sql = "select CustID,CustNum,Name,Address1,Address2,Address3,City,State,Zip,Customer.checkbox01,Terms.Description as TermsDesc,CreditLimit,customer.TermsCode,Customer.Number02 as TCLAmt from Customer left outer join Erp.Terms on Customer.Company = Terms.Company and Terms.TermsCode = customer.TermsCode where CustID='" & _p_parameter.replace("'", "''") & "'"

	if _p_parameter2 <> "" then 
		sql = sql & " and SalesRepCode = '" & _p_parameter2.replace("'", "''") & "'"
	end if 


            conn = New OdbcConnection(connectionString)
            conn.Open()
            comm = New OdbcCommand(sql, conn)

            ad.SelectCommand = comm
            ad.Fill(ds, "datarecords")
            For Each dr In ds.Tables("datarecords").Rows

                CustName = dr("Name") & ""
                CustID = dr("CustID") & ""
                CustNum = dr("CustNum") & ""
                Address1 = dr("Address1") & ""
                Address2 = dr("Address2") & ""
                Address3 = dr("Address3") & ""
                City = dr("City") & ""
                State = dr("State") & ""
                PostCode = dr("zip") & ""
                TermsDescription = dr("TermsDesc") & ""
                CreditLimit = dr("CreditLimit") & ""
                hasbranch = weblib.bittoboolean(dr("checkbox01") & "")

                Try
                    TermsCode = dr("TermsCode") & ""

                Catch ex As Exception

                End Try

                Try
                    TCLAmount = dr("TCLAmt") & ""

                Catch ex As Exception

                End Try

                counter = counter + 1

            Next
            conn.Close()
            comm.Dispose()
            conn.Dispose()

            If counter = 0 Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function getCreditInfo(ByVal _p_parameter As String) As Boolean

        Dim conn As OdbcConnection
        Dim comm As OdbcCommand
        Dim dr As DataRow
        Dim counter As Integer = 0
        Dim ad As New Odbc.OdbcDataAdapter
        Dim ds As New DataSet()
        Dim connectionString As String
        Dim sql As String

        Try


            connectionString = Weblib.connEpicor
            'sql = "select SUM(PostedInvoiceBalance) as InvoiceUnpaid,sum(orderbalance) as OrderUnpaid,CreditLimit  from v_CLUtilization where CustID='" & _p_parameter.replace("'", "''") & "' group by CreditLimit"
            sql = "select SUM(ShippedAmount) as ShipAmt, SUM(UnpostedInvoiceBalance) as InvoiceUnposted, SUM(PostedInvoiceBalance) as InvoiceUnpaid,sum(orderbalance) as OrderUnpaid,CreditLimit  from v_CLUtilization where CustID='" & _p_parameter.replace("'", "''") & "' group by CreditLimit"


            conn = New OdbcConnection(connectionString)
            conn.Open()
            comm = New OdbcCommand(sql, conn)

            ad.SelectCommand = comm
            ad.Fill(ds, "datarecords")
            For Each dr In ds.Tables("datarecords").Rows
                InvoiceCredit = dr("InvoiceUnpaid") & ""
                OrderCredit = dr("OrderUnpaid") & ""
                CreditLimit = dr("CreditLimit") & ""

                InvoiceCreditUnposted = dr("InvoiceUnposted") & ""
                ShipAmt = dr("ShipAmt") & ""

                counter = counter + 1
                Exit For
            Next
            conn.Close()
            comm.Dispose()
            conn.Dispose()


            If isnumeric(InvoiceCredit) = False Then
                InvoiceCredit = "0.00"
            End If
            If isnumeric(OrderCredit) = False Then
                OrderCredit = "0.00"
            End If
            If isnumeric(InvoiceCreditUnposted) = False Then
                InvoiceCreditUnposted = "0.00"
            End If
            If isnumeric(ShipAmt) = False Then
                ShipAmt = "0.00"
            End If



            'TotalCredit = CDbl(InvoiceCredit) + CDbl(OrderCredit)
            'TotalCredit = CDbl(InvoiceCredit) + CDbl(OrderCredit) + CDbl(InvoiceCreditUnposted) + CDbl(ShipAmt)

            'Dor staff inclusive of order amt
            TotalCredit2 = CDbl(InvoiceCredit) + CDbl(OrderCredit) + CDbl(InvoiceCreditUnposted) + CDbl(ShipAmt)

            'For Customer
            'TotalCredit = CDbl(InvoiceCredit) + CDbl(OrderCredit) + CDbl(InvoiceCreditUnposted) + CDbl(ShipAmt)
            TotalCredit = CDbl(InvoiceCredit) + CDbl(InvoiceCreditUnposted) + CDbl(ShipAmt)

            Try

                If CDbl(TotalCredit) > CDbl(CreditLimit) Then

                    CreditPercent = "100"
                Else
                    If CDbl(CreditLimit) = 0 Then
                        CreditPercent = ""
                    Else
                        CreditPercent = format(((CDbl(TotalCredit) / CDbl(CreditLimit)) * 100), "#0.00")
                    End If
                End If

            Catch ex As Exception
                CreditPercent = ""
            End Try


            Try

                If CDbl(TotalCredit2) > CDbl(CreditLimit) Then

                    CreditPercent2 = "100"
                Else
                    If CDbl(CreditLimit) = 0 Then
                        CreditPercent2 = ""
                    Else
                        CreditPercent2 = format(((CDbl(TotalCredit2) / CDbl(CreditLimit)) * 100), "#0.00")
                    End If
                End If

            Catch ex As Exception
                CreditPercent = ""
            End Try



            If counter = 0 Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function getOpenInvoiceCount(ByVal custnum As String) As Boolean

        Dim conn As OdbcConnection
        Dim comm As OdbcCommand
        Dim dr As DataRow
        Dim counter As Integer = 0
        Dim ad As New Odbc.OdbcDataAdapter
        Dim ds As New DataSet()
        Dim connectionString As String
        Dim sql As String

        Try

            If isnumeric(custnum) = False Then
                OpenInvoice = "-Error-"
                Return False
            End If

            connectionString = Weblib.connEpicor
            sql = "select isnull(count(InvoiceNum),0) as total from invchead where OpenInvoice = 1 and CreditMemo<>1 and custnum=" & custnum

            conn = New OdbcConnection(connectionString)
            conn.Open()
            comm = New OdbcCommand(sql, conn)

            ad.SelectCommand = comm
            ad.Fill(ds, "datarecords")
            For Each dr In ds.Tables("datarecords").Rows
                OpenInvoice = dr("total") & ""
                counter = counter + 1
                Exit For
            Next
            conn.Close()
            comm.Dispose()
            conn.Dispose()

            Return True
        Catch ex As Exception
            OpenInvoice = "-Error-"
            Return False
        End Try
    End Function
    Public Function GetSumAgeingSQL(ByVal CustNum As String)

        If isnumeric(custnum) = False Then
            Return ""
            Exit Function
        End If
        Dim _l_sql As String = ""
        _l_sql = _l_sql & "Select "
        _l_sql = _l_sql & "sum(CASE WHEN DATEDIFF(dd, invchead.invoicedate, GETDATE()) <= 30 THEN InvcHead.DocInvoiceBal ELSE 0 END) AS [0 - 30 Days], "
        _l_sql = _l_sql & "sum(CASE WHEN datediff(dd, invchead.invoicedate, GETDATE()) > 30 AND DATEDIFF(dd, invchead.invoicedate, GETDATE()) <= 60 THEN InvcHead.DocInvoiceBal ELSE 0 END) AS [31 - 60 Days], "
        _l_sql = _l_sql & "sum(CASE WHEN datediff(dd, invchead.invoicedate, GETDATE()) > 60 AND DATEDIFF(dd, invchead.invoicedate, GETDATE()) <= 90 THEN InvcHead.DocInvoiceBal ELSE 0 END) AS [61 - 90 Days], "
        _l_sql = _l_sql & "sum(CASE WHEN datediff(dd, invchead.invoicedate, GETDATE()) > 90 AND DATEDIFF(dd, invchead.invoicedate, GETDATE()) <= 120 THEN InvcHead.DocInvoiceBal ELSE 0 END) AS [91 - 120 Days], "
        _l_sql = _l_sql & "sum(CASE WHEN DATEDIFF(dd, invchead.invoicedate, GETDATE()) > 120 THEN InvcHead.DocInvoiceBal ELSE 0 END) AS [Above 120 Days] "
        _l_sql = _l_sql & " from invchead where custnum=" & CustNum & " and  InvcHead.DocInvoiceBal <> 0"

        Return _l_sql
    End Function


    Public Function isShipToDuplicate(ByVal custnum As String, ByVal ShiptoCode As String) As Boolean

        Dim conn As OdbcConnection
        Dim comm As OdbcCommand
        Dim dr As DataRow
        Dim counter As Integer = 0
        Dim ad As New Odbc.OdbcDataAdapter
        Dim ds As New DataSet()
        Dim connectionString As String
        Dim sql As String

        Try

            If isnumeric(custnum) = False Then
                Return True
            End If

            connectionString = Weblib.connEpicor
            sql = "select shiptonum from shipto where shiptonum ='" & ShiptoCode.trim & "' and custnum=" & custnum

            conn = New OdbcConnection(connectionString)
            conn.Open()
            comm = New OdbcCommand(sql, conn)

            ad.SelectCommand = comm
            ad.Fill(ds, "datarecords")
            For Each dr In ds.Tables("datarecords").Rows
                counter = counter + 1
                Exit For
            Next
            conn.Close()
            comm.Dispose()
            conn.Dispose()

            If counter = 1 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return True
        End Try
    End Function

    Public Function GetBalanceSheetByCustomer_DS(ByVal custcode As String, ByVal topN As Integer, ByVal fromtopyear As Integer) As dataset
        Dim conn As OdbcConnection
        Dim comm As OdbcCommand
        Dim dr As DataRow
        Dim counter As Integer = 0
        Dim ad As New Odbc.OdbcDataAdapter
        Dim ds As New DataSet()
        Dim connectionString As String
        Dim sql As String


        Try

            connectionString = Weblib.connEpicor
            'sql = "select top " & TopN & " Key2 as custcode,Key3 as [datayear], 0 as currentasset, 0 as currentliabilities, 0 as nettworkingcapital,0 as otherassets,0 as paidupcapital, 0 as networth,0 as longtermdebts, 0 as tradecreditors, number09 as pnlsales,Number10 as pnlprofitbeforetax, number13 as ratiocurrent,number14 as ratiodebt,number18 as ratiocreditor from UD04 where Key2='" & custcode & "' and key3<='" & fromtopyear & "' order by Key3 desc"
            'sql = "select top " & TopN & " Key2 as custcode,Key3 as [datayear],date01 as datadate, Number24 as currentasset, Number25 as currentliabilities, Number21 as nettworkingcapital,Number26 as otherassets,Number27 as paidupcapital, Number28 as networth,Number22 as longtermdebts, Number29 as tradecreditors, number09 as pnlsales,Number10 as pnlprofitbeforetax, number13 as ratiocurrent,number14 as ratiodebt,number18 as ratiocreditor from UD04 inner join UD04_ud on UD04.sysrowid = UD04_ud.foreignsysrowid where Key2='" & custcode & "' and key3<='" & fromtopyear & "' order by Key3 desc"
            'sql = "select top " & TopN & " Key2 as custcode,Key3 as [datayear],date01 as datadate, Number24 as currentasset, Number25 as currentliabilities, Number21 as nettworkingcapital,Number26 as otherassets,Number27 as paidupcapital, Number28 as networth,Number22 as longtermdebts, Number29 as tradecreditors, number09 as pnlsales,Number10 as pnlprofitbeforetax, number13 as ratiocurrent,number14 as ratiodebt,number18 as ratiocreditor,number30 as DSO from UD04 inner join UD04_ud on UD04.sysrowid = UD04_ud.foreignsysrowid where Key2='" & custcode & "' and key3<='" & fromtopyear & "' order by Key3 desc"
            sql = "select top " & TopN & " Key2 as custcode,Key3 as [datayear],date01 as datadate, Number24 as currentasset, Number25 as currentliabilities, Number21 as nettworkingcapital,Number26 as otherassets,Number27 as paidupcapital, Number28 as networth,Number22 as longtermdebts, Number29 as tradecreditors, number09 as pnlsales,Number10 as pnlprofitbeforetax, number13 as ratiocurrent,number14 as ratiodebt,number18 as ratiocreditor,number01 as DSO from Ice.UD04 inner join Ice.UD04_ud on UD04.sysrowid = UD04_ud.foreignsysrowid where Key2='" & custcode & "' and key3<='" & fromtopyear & "' order by Key3 desc"

            'select top 3 Key2 as custcode,Key3 as [datayear], Number24 as currentasset, Number25 as currentliabilities, Number21 as nettworkingcapital,Number26 as otherassets,Number27 as paidupcapital, Number28 as networth,Number22 as longtermdebts, Number29 as tradecreditors, number09 as pnlsales,Number10 as pnlprofitbeforetax, number13 as ratiocurrent,number14 as ratiodebt,number18 as ratiocreditor from UD04 inner join UD04_ud on UD04.sysrowid = UD04_ud.foreignsysrowid where Key2='ABC' and key3<='2016' order by Key3 desc

            conn = New OdbcConnection(connectionString)
            conn.Open()
            comm = New OdbcCommand(sql, conn)

            ad.SelectCommand = comm
            ad.Fill(ds, "datarecords")

            Return ds

            conn.Close()
            comm.Dispose()
            conn.Dispose()


        Catch ex As Exception
            Return Nothing
        End Try




    End Function


End Class

Public Class RuntimeShipTo
    Public ShipToNum As String
    Public ShipToName As String
    Public ShipToAddress1 As String
    Public ShipToAddress2 As String
    Public ShipToAddress3 As String
    Public ShipToAddress4 As String
    Public ShipToCity As String
    Public ShipToState As String
    Public ShipToPostCode As String
    Public ShipToCountry As String
    Public Function getInfo(ByVal ShipToNum As String, ByVal CustNum As String) As Boolean


        Dim conn As OdbcConnection
        Dim comm As OdbcCommand
        Dim dr As DataRow
        Dim counter As Integer = 0
        Dim ad As New Odbc.OdbcDataAdapter
        Dim ds As New DataSet()
        Dim connectionString As String
        Dim sql As String

        Try

            connectionString = Weblib.connEpicor

            sql = "select shiptonum,name as shiptoname,Address1,Address2, Address3,City ,State ,Zip,Country  from ShipTo where shiptonum ='" & ShipToNum & "' and CustNum=" & CustNum

            conn = New OdbcConnection(connectionString)
            conn.Open()
            comm = New OdbcCommand(sql, conn)

            ad.SelectCommand = comm
            ad.Fill(ds, "datarecords")
            For Each dr In ds.Tables("datarecords").Rows
                shiptonum = dr("shiptonum") & ""
                ShipToName = dr("shiptoname") & ""
                ShipToAddress1 = dr("Address1") & ""
                ShipToAddress2 = dr("Address2") & ""
                ShipToAddress3 = dr("Address3") & ""
                ShipToAddress4 = "" & ""
                ShipToCity = dr("City") & ""
                ShipToState = dr("State") & ""
                ShipToPostCode = dr("Zip") & ""
                ShipToCountry = dr("Country") & ""
                counter = counter + 1
                Exit For
            Next
            conn.Close()
            comm.Dispose()
            conn.Dispose()

            If counter = 0 Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function
End Class
Public Class RuntimeDestinationCode
    Public Description As String
    Public Code As String
    Public Key3 As String
    Public Function getInfo(ByVal _p_parameter As String) As Boolean


        Dim conn As OdbcConnection
        Dim comm As OdbcCommand
        Dim dr As DataRow
        Dim counter As Integer = 0
        Dim ad As New Odbc.OdbcDataAdapter
        Dim ds As New DataSet()
        Dim connectionString As String
        Dim sql As String

        Try


            connectionString = Weblib.connEpicor
            sql = "Select character01,key1,key3 from Ice.UD100 where key1='" & _p_parameter.replace("'", "''") & "'"

            conn = New OdbcConnection(connectionString)
            conn.Open()
            comm = New OdbcCommand(sql, conn)

            ad.SelectCommand = comm
            ad.Fill(ds, "datarecords")
            For Each dr In ds.Tables("datarecords").Rows

                Description = dr("character01") & ""
                Code = dr("key1") & ""
                Key3 = dr("key3") & ""

                counter = counter + 1

            Next
            conn.Close()
            comm.Dispose()
            conn.Dispose()

            If counter = 0 Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

End Class
Public Class RuntimeCustomerBranch
    Public Description As String
    Public Code As String
    Public Function getInfo(ByVal p_ContactNum As String, ByVal p_custnum As String) As Boolean


        Dim conn As OdbcConnection
        Dim comm As OdbcCommand
        Dim dr As DataRow
        Dim counter As Integer = 0
        Dim ad As New Odbc.OdbcDataAdapter
        Dim ds As New DataSet()
        Dim connectionString As String
        Dim sql As String

        Try


            connectionString = Weblib.connEpicor
            sql = "Select shortchar02,shortchar03 from custcnt where custcnt.ConNum=" & p_ContactNum & " and custcnt.CustNum =" & p_custnum & " and custcnt.SpecialAddress=1"
            'shortchar02 code
            'shortchar03 name
            conn = New OdbcConnection(connectionString)
            conn.Open()
            comm = New OdbcCommand(sql, conn)

            ad.SelectCommand = comm
            ad.Fill(ds, "datarecords")
            For Each dr In ds.Tables("datarecords").Rows

                Description = dr("shortchar03") & ""
                Code = dr("shortchar02") & ""

                counter = counter + 1
                Exit For
            Next
            conn.Close()
            comm.Dispose()
            conn.Dispose()

            If counter = 0 Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

End Class