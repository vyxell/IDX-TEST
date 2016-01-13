Imports AAS.DataAccess

Public Class Offering : Inherits BusinessBase

    Function GetListOfOffering(Optional ByVal parOfferingID As String = "") As DataTable
        With New DataAccess.Offering()
            Return .GetListOfOffering(parOfferingID)
        End With
    End Function

    Function GetSelectionOfOffering() As DataTable
        With New DataAccess.Offering()
            Return .GetSelectionOfOffering()
        End With
    End Function

    Function InsertOffering(ByVal parOffering As AAS.Business.Entity.Offering, _
            ByVal parMode As FormMode, _
            ByVal parUserLogin As String _
    ) As Boolean
        With New DataAccess.Offering()
            Try
                Return .InsertOffering(parOffering, parMode, parUserLogin)
            Catch ex As Exception
                Throw ex
            End Try
        End With
    End Function

    Function DeleteOffering(ByVal parOfferingID As String, _
            ByVal parMode As FormMode, _
            ByVal parUserLogin As String _
    ) As Boolean
        With New DataAccess.Offering()
            Try
                Return .DeleteOffering(parOfferingID, parMode, parUserLogin)
            Catch ex As Exception
                Throw ex
            End Try
        End With
    End Function

    Function CekOfferingExsist(ByVal parOffID As String)
        With New DataAccess.Offering()
            Try
                Return .CekOfferingExsist(parOffID)
            Catch ex As Exception
                Throw ex
            End Try
        End With
    End Function
End Class
