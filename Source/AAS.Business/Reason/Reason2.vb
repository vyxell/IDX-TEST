Imports AAS.DataAccess

Public Class Reason2 : Inherits BusinessBase

    Function GetListOfReason2(Optional ByVal parReason2ID As String = "") As DataTable
        With New DataAccess.Reason2()
            Return .GetListOfReason2(parReason2ID)
        End With
    End Function

    Function CekReason2Exsist(ByVal parReason2ID As String) As Boolean
        With New DataAccess.Reason2()
            Try
                Return .CekReason2Exsist(parReason2ID)
            Catch ex As Exception
                Throw ex
            End Try
        End With
    End Function

    Function GetSelectionOfReason2(ByVal parReason1ID As String) As DataTable
        With New DataAccess.Reason2()
            Return .GetSelectionOfReason2(parReason1ID)
        End With
    End Function

    Function InsertReason2(ByVal parReason2 As AAS.Business.Entity.Reason2, _
            ByVal parMode As FormMode, _
            ByVal parUserLogin As String _
    ) As Boolean
        With New DataAccess.Reason2()
            Return .InsertReason2(parReason2, parMode, parUserLogin)
        End With
    End Function

    Function DeleteReason2(ByVal parReason2ID As String, _
            ByVal parMode As FormMode, _
            ByVal parUserLogin As String _
    ) As Boolean
        With New DataAccess.Reason2()
            Return .DeleteReason2(parReason2ID, parMode, parUserLogin)
        End With
    End Function

End Class
