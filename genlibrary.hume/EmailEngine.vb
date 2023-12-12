Imports Microsoft.VisualBasic

Public Class clsEmailNoti
    Public Function GetEmailCustomFiels(ByVal appmodule As String)

        Select Case appmodule.tostring.tolower
            Case "workflow"
                Return CustomFieldsWorkflow

        End Select

    End Function

    Public ReadOnly Property CustomFieldsWorkflow()
        Get
            Return "#WorkFlowName#|#SecureURL#|#DocumentNo#"
        End Get
    End Property

End Class
