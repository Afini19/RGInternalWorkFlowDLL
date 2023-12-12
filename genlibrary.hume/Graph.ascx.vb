Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Partial Public Class graph_class
    Inherits System.Web.UI.UserControl
    Protected _Width As String = "300px"
    Protected _Height As String = "600px"
    Protected _ChartTitle As String = ""
    Protected _ChartSubTitle As String = ""
    Protected _ToolTipValueSuffix As String = ""
    Protected _XAxisTitle As String = ""
    Protected _YAxisTitle As String = ""
    Protected _ChartType As String = ""
    Protected _YAxisMinValue As String = "0"
    Protected _XAxisLabels As String = ""
    Protected _PiePErcent As Boolean = True

    Public litcode, graphcontainer As Object
    Protected _Data As String = ""

    Public Property Width() As String
        Get
            Return _Width
        End Get
        Set(ByVal value As String)
            _Width = value
        End Set
    End Property
    Public Property Height() As String
        Get
            Return _Height
        End Get
        Set(ByVal value As String)
            _Height = value
        End Set
    End Property
    Public Property ChartType() As String
        Get
            Return _ChartType
        End Get
        Set(ByVal value As String)
            _ChartType = value
        End Set
    End Property
    Public Property Data() As String
        Get
            Return _Data
        End Get
        Set(ByVal value As String)
            _Data = value
        End Set
    End Property
    Public Property PieinPercentage() As Boolean
        Get
            Return _PiePErcent
        End Get
        Set(ByVal value As Boolean)
            _PiePErcent = value
        End Set
    End Property

    Public Property ChartTitle() As String
        Get
            Return "title: {text:'" & _ChartTitle & "'}"
        End Get
        Set(ByVal value As String)
            _ChartTitle = value
        End Set
    End Property
    Public Property ChartSubTitle() As String
        Get
            Return "subtitle: {text: '" & _ChartSubTitle & "'}"
        End Get
        Set(ByVal value As String)
            _ChartSubTitle = value
        End Set
    End Property
    Public Property XAxisTitle() As String
        Get
            Return _XAxisTitle
        End Get
        Set(ByVal value As String)
            _XAxisTitle = value
        End Set
    End Property
    Public Property XAxisLabels() As String
        Get
            Return _XAxisLabels
        End Get
        Set(ByVal value As String)
            _XAxisLabels = value
        End Set
    End Property

    Public Property YAxisTitle() As String
        Get
            Return _YAxisTitle
        End Get
        Set(ByVal value As String)
            _YAxisTitle = value
        End Set
    End Property
    Public Property YAxisMinValue() As String
        Get
            If isnumeric(_YAxisMinValue) = False Then
                _YAxisMinValue = "0"
            End If
            Return _YAxisMinValue
        End Get
        Set(ByVal value As String)
            _YAxisMinValue = value
        End Set
    End Property
    Public Property ToolTipValueSuffix() As String
        Get
            Return _ToolTipValueSuffix
        End Get
        Set(ByVal value As String)
            _ToolTipValueSuffix = value
        End Set
    End Property
    Private Function Chart() As String
        Dim ltemp As String = ""

        Select Case _ChartType.ToLower

            Case "bar"
                ltemp = "chart: {type: 'bar'}"

            Case "pie"
                ltemp = "chart: {plotBackgroundColor: null,plotBorderWidth: null,plotShadow: false}"

            Case "line"
                ltemp = "chart: {type: 'line'}"

        End Select
        Return ltemp


    End Function
    Private Function ToolTip() As String
        Dim ltemp As String = ""

        Select Case _ChartType.ToLower

            Case "bar"
                ltemp = "valueSuffix: '" & ToolTipValueSuffix & "'"

            Case "pie"
                If PieinPercentage = True Then
                    ltemp = "pointFormat: '{series.name}: <b>{point.percentage:.1f}%</b>'"
                End If
            Case "line"
                ltemp = "valueSuffix: '" & ToolTipValueSuffix & "'"

        End Select
        Return "tooltip: {" & ltemp & "}"

    End Function
    Private Function XAxis() As String
        Dim ltemp As String = ""

        Select Case _ChartType.ToLower

            Case "bar"
                ltemp = "categories: [" & XAxisLabels & "],title: {text: '" & XAxisTitle & "'}"

            Case "pie"

            Case "line"
                ltemp = "categories: [" & XAxisLabels & "],title: {text: '" & XAxisTitle & "'}"

        End Select
        Return "xAxis: {" & ltemp & " }"

    End Function
    Private Function YAxis() As String
        Dim ltemp As String = ""

        Select Case _ChartType.ToLower

            Case "bar"
                ltemp = " min: " & YAxisMinValue & ",title: {text: '" & YAxisTitle & "',align: 'high'},labels: {overflow:  'justify'}"

            Case "pie"

            Case "line"
                ltemp = " min: " & YAxisMinValue & ",title: {text: '" & YAxisTitle & "',align: 'high'},labels: {overflow:  'justify'}"

        End Select
        Return "yAxis: {" & ltemp & "}"


    End Function
    Private Function Series() As String
        Dim ltemp As String = ""
        Dim tempo
        Dim bSingleSeries As Boolean = False
        If ChartType.ToLower = "pie" Then
            bSingleSeries = True
        End If

        If _Data.Trim = "" Then
            Return "'"
            Exit Function
        End If

        tempo = Microsoft.VisualBasic.Split(_Data, ";;")

        For counter = 0 To Microsoft.VisualBasic.UBound(tempo)
            Dim ltempa

            ltempa = Microsoft.VisualBasic.Split(tempo(counter), "||")

            If Microsoft.VisualBasic.UBound(ltempa) < 0 Then
                GoTo NExtItem
            End If

            If ltemp.Trim <> "" Then
                ltemp = ltemp & ","
            End If





            Select Case _ChartType.ToLower

                Case "bar", "line"
                    ltemp = ltemp & "{name:'" & ltempa(0) & "',data: [" & ltempa(1) & "]}"

                Case "pie"
                    ltemp = ltemp & "{type:'pie',name:'" & ltempa(0) & "',data: [" & ltempa(1) & "]}"

                Case Else

                    ltemp = ltemp & ""
            End Select


            '            ltemp = ""

            '                Case "pie"
            '           ltemp = ""

            '              Case "line"
            '         ltemp = ""


            '            End Select


NextITem:

            If bSingleSeries = True And ltemp.Trim <> "" Then
                Exit For
            End If

        Next
        Return "series: [" & ltemp & "]"



        '        Return "series: [{name:'Target (Tonnes)',data: [107, 331, 635, 203, 21,324,213,455,211,789,321,445]}, {name:   'Actual (Tonnes)',data: [973, 914, 4054, 732, 314, 321,467,764,234,678,88,332]}]"



    End Function
    Private Function Legend() As String
        Dim ltemp As String = ""

        Select Case _ChartType.ToLower

            Case "bar"
                ltemp = "layout:'vertical',align: 'right',verticalAlign: 'top',x: -40,y: 100,floating: true,borderWidth: 1,backgroundColor: (Highcharts.theme && Highcharts.theme.legendBackgroundColor || '#FFFFFF'),shadow: true"


            Case "pie"

            Case "line"
                ltemp = "layout:'horizontal',align: 'center',verticalAlign: 'top',x: -40,y: 100,borderWidth: 1,backgroundColor: (Highcharts.theme && Highcharts.theme.legendBackgroundColor || '#FFFFFF'),shadow: true"

        End Select
        Return " legend: {" & ltemp & "}"

    End Function
    Private Function PlotOptions() As String
        Dim ltemp As String = ""
        Select Case _ChartType.ToLower

            Case "bar"
                ltemp = "bar: { dataLabels: {enabled: true}}"
            Case "pie"
                ltemp = "pie: {allowPointSelect: true,cursor:'pointer',dataLabels: {enabled: true"
                If PieinPercentage = True Then
                    ltemp = ltemp & " ,format: '<b>{point.name}</b>: {point.percentage:.1f} %'"
                Else
                    '  ltemp = ltemp & " ,format: '<b>{point.name}</b>: {point.percentage:.1f} %'"
                End If
                ltemp = ltemp & ", showInLegend: true,style: {color: (Highcharts.theme && Highcharts.theme.contrastTextColor) || 'black'}}}"



            Case "line"
                ltemp = "line: { dataLabels: {enabled: true}}"



        End Select
        Return "plotOptions: {" & ltemp & "}"
    End Function
    Private Function GenerateGraph() As String
        Dim ltemp As String = ""
        ltemp = ltemp & "$(function () {" & environment.NewLine
        ltemp = ltemp & " $('#" & graphcontainer.ClientID & "').highcharts({" & environment.NewLine
        Select Case _ChartType.ToLower

            Case "bar"
                ltemp = ltemp & Chart() & "," & ChartTitle & "," & ChartSubTitle & "," & XAxis() & "," & YAxis() & "," & ToolTip() & "," & PlotOptions() & "," & Legend() & "," & Series()

            Case "pie"
                ltemp = ltemp & Chart() & "," & ChartTitle & "," & ChartSubTitle & "," & ToolTip() & "," & PlotOptions() & "," & Legend() & "," & Series()
            Case "line"
                ltemp = ltemp & Chart() & "," & ChartTitle & "," & ChartSubTitle & "," & XAxis() & "," & YAxis() & "," & ToolTip() & "," & PlotOptions() & "," & Legend() & "," & Series()

        End Select

        ltemp = ltemp & "});});"
        Return ltemp

    End Function
    Public Sub InitGraph()
        litcode.text = "<script>" & GenerateGraph() & "</script>"
    End Sub
End Class

