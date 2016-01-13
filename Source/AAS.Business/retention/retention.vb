Imports AAS.DataAccess

Public Class retention : Inherits BusinessBase

    Function CreateSPRetention()
        With New AAS.DataAccess.retention
            .CreateSPRetention()
        End With
    End Function

    Function GetListOfRetention(Optional ByVal parRttCode As String = "") As DataTable
        With New DataAccess.retention()
            Return .GetListOfRetention(parRttCode)
        End With
    End Function

    Function GetListOfRetention_Offer(Optional ByVal parRttCode As String = "") As DataTable
        With New DataAccess.retention()
            Return .GetListOfRetention_Offer(parRttCode)
        End With
    End Function

    Function InsertRetention(ByVal parData As AAS.Business.Entity.retention, _
            ByVal parFinalOffer As ArrayList, _
            ByVal parMode As FormMode, _
            ByVal parUserLogin As String _
    ) As Boolean
        With New DataAccess.retention()
            Try
                Return .InsertRetention(parData, parFinalOffer, parMode, parUserLogin)
            Catch ex As Exception
                Throw ex
            End Try
        End With
    End Function

    Function ValidateRule(ByVal parRule As String, ByRef parMsg As String) As Boolean
        With New DataAccess.retention()
            Try
                Return .ValidateRule(parRule)
            Catch ex As Exception
                parMsg = ex.Message.Replace("'", "\'").Replace(vbCrLf, "\n")
                Return False
            End Try
        End With
    End Function

    Function DeleteRetention(ByVal parRettCode As String, _
        ByVal parMode As FormMode, _
        ByVal parUserLogin As String _
) As Boolean
        With New DataAccess.retention()
            Try
                Return .DeleteRetention(parRettCode, parMode, parUserLogin)
            Catch ex As Exception
                Throw ex
            End Try
        End With
    End Function

    Function CekRetentionExsist(ByVal parRettCode As String)
        With New DataAccess.retention()
            Try
                Return .CekRetentionExsist(parRettCode)
            Catch ex As Exception
                Throw ex
            End Try
        End With
    End Function

    Public Function GetRuleField() As DataTable
        With New AAS.DataAccess.retention()
            Return .GetRuleField()
        End With
    End Function
End Class