Public Class etl
    Function run(ByVal parUserID As String) As Boolean
        Try
            With New AAS.DataAccess.scorecard()
                .CreateSPScorecard()
            End With
            With New AAS.DataAccess.retention()
                .CreateSPRetention()
            End With
            With New AAS.DataAccess.Systems()
                .RunETL(parUserID)
            End With
        Catch ex As Exception
            Throw ex
        End Try
    End Function
End Class
