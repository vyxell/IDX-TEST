Imports AAS.DataAccess

Public Class Level : Inherits BusinessBase
    Function GetListOfLevel(Optional ByVal paramLevelID As String = "") As DataTable
        With New DataAccess.Level()            
            Return .GetListOfLevel(paramLevelID)
        End With
    End Function

    Function GetAccessUserMenu(ByVal parUserId As String, ByVal parMenu As String) As DataTable
        With New DataAccess.Level()
            Return .GetAccessUserMenu(parUserId, parMenu)
        End With
    End Function

    Function GetListOfLevelMenu(Optional ByVal paramLevelID As String = "") As DataTable
        With New DataAccess.Level()
            Return .GetListOfLevelMenu(paramLevelID)
        End With
    End Function

    Function GetLevelDetail(Optional ByVal paramLevelID As String = "") As DataTable
        With New DataAccess.Level()
            Return .GetLevelDetail(paramLevelID)
        End With
    End Function

    Function DeleteLevel(ByVal paramLevelID As String, _
            ByVal paramMode As FormMode, _
            ByVal paramUserLogin As String _
    ) As Boolean
        With New DataAccess.Level()
            Try
                Return .DeleteLevel(paramLevelID, paramMode, paramUserLogin)
            Catch ex As Exception
                Throw ex
            End Try
        End With
    End Function

    Function InsertLevel(ByVal parLevel As AAS.Business.Entity.Level, _
            ByVal paramMode As FormMode, _
            ByVal paramUserLogin As String _
    ) As Boolean
        With New DataAccess.Level()
            Try
                Return .InsertLevel(parLevel, paramMode, paramUserLogin)
            Catch ex As Exception
                Throw ex
            End Try
        End With
    End Function

    Function CekLevelExsist(ByVal paramLevelID As String) As Boolean
        With New DataAccess.Level()
            Try
                Return .CekLevelExsist(paramLevelID)
            Catch ex As Exception
                Throw ex
            End Try
        End With
    End Function
End Class
