Public Class query : Inherits BusinessBase

    Function GetListOfQuery(Optional ByVal parQueryID As String = "") As DataTable
        With New DataAccess.query()
            Return .GetListOfQuery(parQueryID)
        End With
    End Function

    Function InsertQuery(ByVal parQuery As AAS.Business.Entity.Query, _
            ByVal parMode As FormMode, _
            ByVal parUserLogin As String _
    ) As Boolean
        With New DataAccess.query()
            Return .InsertQuery(parQuery, parMode, parUserLogin)
        End With
    End Function

    Function DeleteQuery(ByVal parQueryID As String, _
            ByVal parMode As FormMode, _
            ByVal parUserLogin As String _
    ) As Boolean
        With New DataAccess.query()
            Return .DeleteQuery(parQueryID, parMode, parUserLogin)
        End With
    End Function

    Function GetTableField(ByVal parTable As String) As DataTable
        With New DataAccess.query()
            Return .GetTableField(parTable)
        End With
    End Function

    Function ValidateField(ByVal parField As String, ByVal parTable As String, ByRef parMsg As String) As Boolean
        With New DataAccess.query()
            Try
                Return .ValidateField(parField, parTable)
            Catch ex As Exception
                parMsg = ex.Message.Replace("'", "\'").Replace(vbCrLf, "\n")
                Return False
            End Try
        End With
    End Function

    Function ValidateFilter(ByVal parFilter As String, ByVal parTable As String, ByRef parMsg As String) As Boolean
        With New DataAccess.query()
            Try
                Return .ValidateFilter(parFilter, parTable)
            Catch ex As Exception
                parMsg = ex.Message.Replace("'", "\'").Replace(vbCrLf, "\n")
                Return False
            End Try
        End With
    End Function

    Function GetListOfExecute(ByVal parQueryID As Int64) As DataTable
        With New DataAccess.query()
            Return .GetListOfExecute(parQueryID)
        End With
    End Function
End Class