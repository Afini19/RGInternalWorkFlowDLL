Imports System.IO

Public Class VIFEANDI_APP_INI
#Region "API Calls"
    Private Declare Unicode Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringW" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, ByVal lpString As String, _
    ByVal lpFileName As String) As Int32

    Private Declare Unicode Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringW" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Int32, _
    ByVal lpFileName As String) As Int32
#End Region


    Public Overloads Function INIRead(ByVal INIPath As String, _
ByVal SectionName As String, ByVal KeyName As String, _
ByVal DefaultValue As String) As String
        ' primary version of call gets single value given all parameters
        Call getdefaultfile(INIPATH)

        Dim n As Int32
        Dim sData As String
        sData = Space$(1024) ' allocate some room 
        n = GetPrivateProfileString(SectionName, KeyName, DefaultValue, _
        sData, sData.Length, INIPath)
        If n > 0 Then ' return whatever it gave us
            INIRead = sData.Substring(0, n)
        Else
            INIRead = ""
        End If
    End Function
    Private Sub getdefaultfile(ByVal INIPath As String)
        Dim Sr As StreamWriter

        Try

            If Not File.Exists(INIPath) Then
                SR = File.CreateText(INIPath)
                SR.Close()
            End If

        Catch ex As Exception
            Throw New System.Exception(ex.Message)

        End Try
    End Sub
    Public Overloads Function INIRead(ByVal INIPath As String, _
    ByVal SectionName As String, ByVal KeyName As String) As String
        ' overload 1 assumes zero-length default
        Call getdefaultfile(INIPATH)
        Return INIRead(INIPath, SectionName, KeyName, "")
    End Function
    Public Overloads Function INIRead(ByVal INIPath As String, _
    ByVal SectionName As String) As String
        ' overload 2 returns all keys in a given section of the given file
        Call getdefaultfile(INIPATH)
        Return INIRead(INIPath, SectionName, Nothing, "")
    End Function
    Public Overloads Function INIRead(ByVal INIPath As String) As String
        ' overload 3 returns all section names given just path
        Call getdefaultfile(INIPATH)

        Return INIRead(INIPath, Nothing, Nothing, "")
    End Function
    Public Sub INIWrite(ByVal INIPath As String, ByVal SectionName As String, _
    ByVal KeyName As String, ByVal TheValue As String)
        Call getdefaultfile(INIPATH)

        Call WritePrivateProfileString(SectionName, KeyName, TheValue, INIPath)
    End Sub
    Public Overloads Sub INIDelete(ByVal INIPath As String, ByVal SectionName As String, _
    ByVal KeyName As String) ' delete single line from section
        Call getdefaultfile(INIPATH)

        Call WritePrivateProfileString(SectionName, KeyName, Nothing, INIPath)
    End Sub
    Public Overloads Sub INIDelete(ByVal INIPath As String, ByVal SectionName As String)
        ' delete section from INI file
        Call getdefaultfile(INIPATH)

        Call WritePrivateProfileString(SectionName, Nothing, Nothing, INIPath)
    End Sub
End Class

