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

Public Class clsFileTypes
    Public FileExtension As String = ""
    Public FileName As String = ""
    Public FileType As String = ""
    Private FileImage As String = "unk64.png"

    Public Property FileImageFile()
        Get
            Return "graphics/filetypes/" & FileImage
        End Get
        Set(ByVal value)


        End Set
    End Property
    Public Function InitFile(ByVal _p_FileName As String) As Boolean

        If _p_FileName.trim = "" Then
            Call ResetFile()
        End If
        FileExtension = System.IO.Path.GetExtension(_p_FileName)
        FileName = _p_FileName
        FileExtension = FileExtension.replace(".", "")
        Call GetFileType(FileExtension)
    End Function
    Private Function GetFileType(ByVal _p_Ext As String) As String

        Select Case _p_Ext.tolower
            Case "doc"
                FileType = "Microsoft Word"
                FileImage = "doc64.png"
            Case "docx"
                FileType = "Microsoft Word"
                FileImage = "docx64.png"

            Case "png"
                FileType = "Png Image"
                FileImage = "png64.png"

            Case "jpg"
                FileType = "Png Image"
                FileImage = "jpg64.png"

            Case "gif"
                FileType = "Gif Image"
                FileImage = "gif64.png"

            Case "ppt"
                FileType = "Microsoft Powerpoint"
                FileImage = "ppt64.png"

            Case "pptx"
                FileType = "Microsoft Powerpoint"
                FileImage = "pptx64.png"

            Case "xls"
                FileType = "Microsoft Excel"
                FileImage = "xls64.png"

            Case "xlsx"
                FileType = "Microsoft Excel"
                FileImage = "xlsx64.png"

            Case "tif"
                FileType = "Tif File"
                FileImage = "tif64.png"

            Case "tiff"
                FileType = "Tiff File"
                FileImage = "tiff64.png"

            Case "pdf"
                FileType = "Adobe PDF File"
                FileImage = "pdf64.png"

            Case Else
                FileType = _p_Ext & " File"
                FileImage = "oth64.png"

        End Select
    End Function
    Private Function ResetFile()
        FileImage = "unk64.png"
        FileType = "No File Uploaded"
        FileExtension = ""
        FileName = ""
    End Function
End Class