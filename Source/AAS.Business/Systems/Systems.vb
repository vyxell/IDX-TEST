
Imports AAS.DataAccess

Public Class Systems : Inherits BusinessBase
    Function GetSystemSetting() As DataTable
        With New DataAccess.Systems()
            Return .GetSystemSetting()
        End With
    End Function

    Function UpdateSystemSetting(ByVal parSystems As AAS.Business.Entity.Systems, _
            ByVal paramMode As FormMode, _
            ByVal paramUserLogin As String _
) As Boolean
        With New DataAccess.Systems()
            Try
                .UpdateSystemSetting(parSystems, paramMode, paramUserLogin)
                .WriteIP(parSystems.Systems_DB_Server_Up)
            Catch ex As Exception
                Throw ex
            End Try
        End With
    End Function

    Function GetIP()
        Dim lValue As String
        With New DataAccess.Systems()
            Try
                lValue = .GetIP()
            Catch ex As Exception
                Throw ex
            End Try
        End With
        Return lValue
    End Function

    Sub RunDTSManual(ByVal parUserID As String)
        Try
            With New Meta.Component.Database()
                .Execute("update systems set sys_dts_last_update='" & Now & "'")
            End With
            With New DataAccess.Systems()
                .RunDTSManual(parUserID)
            End With
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
