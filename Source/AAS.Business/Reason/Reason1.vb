Imports AAS.DataAccess

Public Class Reason1 : Inherits BusinessBase

    Function GetListOfReason1(Optional ByVal parReason1ID As String = "") As DataTable
        With New DataAccess.Reason1()
            Return .GetListOfReason1(parReason1ID)
        End With
    End Function

    Function CekReason1Exsist(ByVal parReason1ID As String) As Boolean
        With New DataAccess.Reason1()
            Try
                Return .CekReason1Exsist(parReason1ID)
            Catch ex As Exception
                Throw ex
            End Try
        End With
    End Function

    Function GetSelectionOfReason1() As DataTable
        With New DataAccess.Reason1()
            Return .GetSelectionOfReason1()
        End With
    End Function

    Function InsertReason1(ByVal parReason1 As AAS.Business.Entity.Reason1, _
            ByVal parMode As FormMode, _
            ByVal parUserLogin As String _
    ) As Boolean
        With New DataAccess.Reason1()
            Return .InsertReason1(parReason1, parMode, parUserLogin)
        End With
    End Function

    Function DeleteReason1(ByVal parReason1ID As String, _
            ByVal parMode As FormMode, _
            ByVal parUserLogin As String _
    ) As Boolean
        With New DataAccess.Reason1()
            Return .DeleteReason1(parReason1ID, parMode, parUserLogin)
        End With
    End Function

End Class
