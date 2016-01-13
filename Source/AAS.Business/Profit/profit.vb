Public Class profit : Inherits BusinessBase
    Function GetProfit() As DataTable
        With New DataAccess.Profit()
            Return .GetProfit()
        End With
    End Function

    Function UpdateProfit(ByVal parProfit As AAS.Business.Entity.profit, _
            ByVal paramMode As FormMode, _
            ByVal paramUserLogin As String _
) As Boolean
        With New DataAccess.profit()
            Try
                .UpdateProfit(parProfit, paramMode, paramUserLogin)
            Catch ex As Exception
                Throw ex
            End Try
        End With
    End Function

    Function ValidateFormula(ByVal parFormula As String, ByRef parMsg As String) As Boolean
        With New DataAccess.profit()
            Try
                Return .ValidateFormula(parFormula)
            Catch ex As Exception
                parMsg = ex.Message.Replace("'", "\'").Replace(vbCrLf, "\n")
                Return False
            End Try
        End With
    End Function

    Public Function GetProfitField() As DataTable
        With New AAS.DataAccess.profit()
            Return .GetProfitField()
        End With
    End Function
End Class
