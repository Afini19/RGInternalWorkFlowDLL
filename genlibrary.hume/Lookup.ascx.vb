Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Partial Public Class lookup_class
    Inherits System.Web.UI.UserControl
    Protected _Width As String = "150px"
    Protected _Height As String = "20px"
    Protected _AdditionalRtn As String = ""
    Protected _AdditionalRtnFields As String = ""
    Protected _Param1 As String = ""
    Protected _LockLookupField As Boolean = False
    Public lookupcode, lookupdefault, litdisplay, lookuphiddenid, Literal1, isLookup, lnkLookup, LookupPage As Object

    Public Property text() As String
        Get
            Return lookupcode.text
        End Get
        Set(ByVal value As String)
            lookupcode.text = value
            lookupdefault.value = value
        End Set
    End Property
    Public Property LockLookupField() As Boolean
        Get
            Return _LockLookupField
        End Get
        Set(ByVal value As Boolean)
            _LockLookupField = value
        End Set
    End Property
    Public Property Parameter1() As String
        Get
            Return _Param1 & ""
        End Get
        Set(ByVal value As String)
            _Param1 = value
        End Set
    End Property


    Public Property AdditioanalParam() As String
        Get
            Return _AdditionalRtn
        End Get
        Set(ByVal value As String)
            _AdditionalRtn = value
        End Set
    End Property

    Public Property textdesc() As String
        Get
            Return litdisplay.text
        End Get
        Set(ByVal value As String)
            litdisplay.text = value
        End Set
    End Property
    Public Property cssclass() As String
        Get
            Return lookupcode.cssclass
        End Get
        Set(ByVal value As String)
            lookupcode.cssclass = value
        End Set
    End Property
    Public Property Enabled() As Boolean
        Get
            Return lookupcode.enabled

        End Get
        Set(ByVal value As Boolean)
            lookupcode.enabled = value
            lnklookup.visible = value
        End Set
    End Property
    Public Property maxlength() As String
        Get
            Return lookupcode.MaxLength

        End Get
        Set(ByVal value As String)
            If isnumeric(value) = True Then
                lookupcode.MaxLength = CInt(value)
            End If
        End Set
    End Property
    Public Property LookupCategory() As String
        Get
            Return LookupPage.value
        End Get
        Set(ByVal value As String)
            LookupPage.value = value
        End Set
    End Property
    Public Property PreviewText() As Boolean
        Get
            Return litdisplay.visible
        End Get
        Set(ByVal value As Boolean)
            litdisplay.visible = value
        End Set
    End Property
    Public Property Width() As String
        Get
            Return _Width
        End Get
        Set(ByVal value As String)
            _Width = value
            lookupcode.Style.Add("width", _Width)
        End Set
    End Property
    Public Property Height() As String
        Get
            Return _Height
        End Get
        Set(ByVal value As String)
            _Height = value
            lookupcode.Style.Add("height", _Height)
            lnklookup.Style.Add("height", _Height)

        End Set
    End Property
    Private Function breakAdditional() As String
        Dim ltemp As String = ""
        If AdditioanalParam.Trim = "" Then
            Return ""
            Exit Function
        End If

        Dim temp

        Try

            temp = Microsoft.VisualBasic.Split(AdditioanalParam, "|")

            If Microsoft.VisualBasic.UBound(temp) <= 0 Then
                ltemp = ltemp & "$('#" & AdditioanalParam & "').val($('#param1').val());" & environment.NEwLine
                Return ltemp
                Exit Function
            End If



            For counter = 0 To Microsoft.VisualBasic.UBound(temp)

                If temp(counter).trim <> "" Then
                    ltemp = ltemp & "$('#" & temp(counter) & "').val($('#param" & counter + 1 & "').val());" & environment.NEwLine
                End If
            Next
            Return ltemp
        Catch ex As Exception
            Return ""
        End Try
    End Function
    Private Function setdisablefield() As String
        Dim ltemp As String = ""
        If AdditioanalParam.Trim = "" Then
            Return ""
            Exit Function
        End If
        If LockLookupField = False Then
            Return ""
            Exit Function
        End If
        Dim temp

        Try
            temp = Microsoft.VisualBasic.Split(AdditioanalParam, "|")

            If Microsoft.VisualBasic.UBound(temp) <= 0 Then
                ltemp = ltemp & "$('#" & AdditioanalParam & "').prop('disabled', ptruefalse);"

            Else

                For counter = 0 To Microsoft.VisualBasic.UBound(temp) - 1
                    'If temp(counter).trim <> "" And temp(counter).trim.tolower() <> "sa_shiptostate" And temp(counter).trim.tolower() <> "sa_cclass" Then   'No choice but to hard code state 'Epicor got no state
                    If temp(counter).trim <> "" And temp(counter).trim.tolower() = "pt_desch" Then
                        ltemp = ltemp & "customenabled2('" & temp(counter) & "',ptruefalse);" & Environment.NewLine

                    ElseIf temp(counter).trim <> "" And temp(counter).trim.tolower() <> "sa_shiptostate" And temp(counter).trim.tolower() <> "sa_cclass" And temp(counter).trim.tolower() <> "pt_desch" Then   'No choice but to hard code state 'Epicor got no state
                      
                        ltemp = ltemp & "$('#" & temp(counter) & "').prop('readonly', ptruefalse);" & environment.NEwLine

                    Else
                        ltemp = ltemp & "customenabled('" & temp(counter) & "',ptruefalse);" & environment.NEwLine

                        ' $("#dropdown").prop("disabled", true);
                        '
                        '                       ltemp = ltemp & "$('#" & temp(counter) & "').prop('background-color', '#DEDEDE');" & environment.NEwLine
                    End If
                Next


            End If

            ltemp = "function " & lookupcode.ClientID & "_disabled" & "(ptruefalse){" & environment.NEwLine & ltemp & environment.NEwLine & "}" & environment.NEwLine

            Return ltemp
        Catch ex As Exception
            Return ""
        End Try
    End Function
    Protected Overrides Sub OnPreRender(ByVal e As System.EventArgs)
        MyBase.OnPreRender(e)



        If LookupPage.value.trim <> "" Then
            Dim _disable As String = ""
            Dim _enable As String = ""

            If LockLookupField = True Then
                _disable = lookupcode.ClientID & "_disabled(true);"
                _enable = lookupcode.ClientID & "_disabled(false);"
            End If



            Dim aa = "if (event.keyCode == 13){" & lookupcode.ClientID & "_lookup();" & "}"
            lookupcode.Attributes.Add("onkeyup", aa)
            lnklookup.Attributes.Add("onclick", lookupcode.ClientID & "_lookup();")

            Dim ltemp As String = ""
            Dim ladditionalpara As String = ""
            ladditionalpara = breakAdditional()

            ltemp = ltemp & "function " & lookupcode.ClientID & "_lookup" & "(){"
            ltemp = ltemp & "$('#" & islookup.ClientID & "').val('Y');" & environment.NEwLine


            Dim lpath As String
            lpath = weblib.ClientURL("lookupgen.aspx")

            '            ltemp = ltemp & "alert('" & lpath & "?lcat=' + $('#" & LookupPage.ClientID & "').val() + '&skey=' + $('#" & lookupcode.ClientID & "').val().replace(' ', '@@SPACE@@') + '&param1=' + $('#" & Parameter1 & "').val()};"
'            ltemp = ltemp & "alert('" & lpath & "?lcat=' + $('#" & LookupPage.ClientID & "').val() +  '&skey=' + $('#" & lookupcode.ClientID & "').val().replace(' ', '@@SPACE@@'));"
            '            ltemp = ltemp & "alert('" & lpath & "?lcat=' + $('#" & LookupPage.ClientID & "').val() + '&skey=' + $('#" & lookupcode.ClientID & "').val().replace(' ', '@@SPACE@@') + '&param1=' + $('#" & Parameter1 & "').val()};"

            If Parameter1.trim <> "" Then
                'ltemp = ltemp & "$.colorbox({opacity:0.5,trapFocus:true,href: function(){return  'lookupgen.aspx?lcat=' + $('#" & LookupPage.ClientID & "').val() + '&skey=' + $('#" & lookupcode.ClientID & "').val() + '&param1=' + $('#" & Parameter1 & "').val();},width:""650px"", height:""500px"","
                '                ltemp = ltemp & "alert('1');"
                'encodeURIComponent
                'ltemp = ltemp & "$.colorbox({opacity:0.5,trapFocus:true,href: function(){return  '" & lpath & "?lcat=' + $('#" & LookupPage.ClientID & "').val() + '&skey=' + $('#" & lookupcode.ClientID & "').val() + '&param1=' + $('#" & Parameter1 & "').val();},width:""650px"", height:""500px"","
                '                ltemp = ltemp & "$.colorbox({opacity:0.5,trapFocus:true,href: function(){return  '" & lpath & "?lcat=' + $('#" & LookupPage.ClientID & "').val() + '&skey=' + $('#" & lookupcode.ClientID & "').val().replace(' ', '@@SPACE@@') + '&param1=' + $('#" & Parameter1 & "').val();},width:""650px"", height:""500px"","
                '                ltemp = ltemp & "$.colorbox({opacity:0.5,trapFocus:true,href: function(){return  '" & lpath & "?lcat=' + $('#" & LookupPage.ClientID & "').val() + '&skey=' + ('sdn').replace(' ', '@@SPACE@@') + '&param1=' + $('#" & Parameter1 & "').val();},width:""650px"", height:""500px"","

                ltemp = ltemp & "$.colorbox({opacity:0.5,trapFocus:true,href: function(){return  '" & lpath & "?lcat=' + $('#" & LookupPage.ClientID & "').val() + '&skey=' + encodeURIComponent($('#" & lookupcode.ClientID & "').val()) + '&param1=' + $('#" & Parameter1 & "').val();},width:""95%"", height:""50%"","
            Else
                '               ltemp = ltemp & "alert('2');"

                'ltemp = ltemp & "$.colorbox({opacity:0.5,trapFocus:true,href: function(){return  'lookupgen.aspx?lcat=' + $('#" & LookupPage.ClientID & "').val() + '&skey=' + $('#" & lookupcode.ClientID & "').val();},width:""650px"", height:""500px"","
                '                ltemp = ltemp & "$.colorbox({opacity:0.5,trapFocus:true,href: function(){return  '" & lpath & "?lcat=' + $('#" & LookupPage.ClientID & "').val() + '&skey=' + $('#" & lookupcode.ClientID & "').val();},width:""650px"", height:""500px"","

                'ltemp = ltemp & "$.colorbox({opacity:0.5,trapFocus:true,href: function(){return  '" & lpath & "?lcat=' + $('#" & LookupPage.ClientID & "').val() + '&skey=' + $('#" & lookupcode.ClientID & "').val();},width:""650px"", height:""500px"","
                'ltemp = ltemp & "$.colorbox({opacity:0.5,trapFocus:true,href: function(){return  '" & lpath & "?lcat=' + $('#" & LookupPage.ClientID & "').val() + '&skey=' + $('#" & lookupcode.ClientID & "').val().replace(' ', '@@SPACE@@');},width:""650px"", height:""500px"","
                ltemp = ltemp & "$.colorbox({opacity:0.5,trapFocus:true,href: function(){return  '" & lpath & "?lcat=' + $('#" & LookupPage.ClientID & "').val() + '&skey=' + encodeURIComponent($('#" & lookupcode.ClientID & "').val());},width:""95%"", height:""50%"","

            End If

            ltemp = ltemp & "onCleanup: function() {" & environment.NEwLine
            ltemp = ltemp & "var returnflag = $('#selectflag').val() + '';" & environment.NEwLine
            ltemp = ltemp & "if (returnflag.length > 0) {" & environment.NEwLine
            ltemp = ltemp & "$('#" & lookupcode.ClientID & "').val($('#returnobj1').val());" & environment.NEwLine
            ltemp = ltemp & "$('#" & lookupdefault.ClientID & "').val($('#returnobj1').val());" & environment.NEwLine
            ltemp = ltemp & "$('#" & litdisplay.ClientID & "').text($('#previewobj').val());" & environment.NEwLine
            ltemp = ltemp & ladditionalpara & environment.NEwLine & _disable
            ltemp = ltemp & "}"
            ltemp = ltemp & "},onClosed: function(){$('#" & lookupcode.ClientID & "').focus();$('#" & islookup.ClientID & "').val('');}"
            ltemp = ltemp & "});   " & environment.NEwLine

            ltemp = ltemp & "}" & environment.NEwLine

            ltemp = ltemp & setdisablefield()


            'Handle KeyEvents
            ltemp = ltemp & "$(function () {" & environment.NEwLine
            ltemp = ltemp & "$(""#" & lookupcode.ClientID & """).blur(function() {" & environment.NEwLine
            ltemp = ltemp & "if (($('#" & lookupcode.ClientID & "').val() != $('#" & lookupdefault.ClientID & "').val()) && ($('#" & lookupcode.ClientID & "').val().length > 0) && ($('#" & islookup.ClientID & "').val() != 'Y')){" & environment.NEwLine
            ltemp = ltemp & "$('#" & lookupcode.ClientID & "').val($('#" & lookupdefault.ClientID & "').val());" & environment.NEwLine
            ltemp = ltemp & "} else { if ($('#" & lookupcode.ClientID & "').val().length == 0){$('#" & lookupdefault.ClientID & "').val('');$('#" & litdisplay.ClientID & "').text('');" & _enable & "}}"
            ltemp = ltemp & "});"

            ltemp = ltemp & "});   " & environment.NEwLine

            Literal1.text = "<script>" & environment.NEwLine & ltemp & environment.NEwLine & "</script>"
        End If
    End Sub
End Class

