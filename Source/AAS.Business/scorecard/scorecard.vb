Imports AAS.DataAccess

Public Class scorecard : Inherits BusinessBase

    Function CreateSPScorecard()
        With New AAS.DataAccess.scorecard
            .CreateSPScorecard()
        End With
    End Function

    Function GetListOfChar(Optional ByVal parSCID As String = "", Optional ByVal parSCHID As Int64 = 0) As DataTable
        With New DataAccess.scorecard()
            Return .GetListOfChar(parSCID, parSCHID)
        End With
    End Function

    Function GetListOfScorecard(Optional ByVal parSCID As String = "") As DataTable
        With New DataAccess.scorecard()
            Return .GetListOfScorecard(parSCID)
        End With
    End Function

    Function InsertSCH(ByVal parData As AAS.Business.Entity.scorecard_char, _
            ByVal parMode As FormMode, _
            ByVal parUserLogin As String _
    ) As Boolean
        With New DataAccess.scorecard()
            Try
                Return .InsertSCH(parData, parMode, parUserLogin)
            Catch ex As Exception
                Throw ex
            End Try
        End With
    End Function

    Function InsertScorecard(ByVal parSCCode As AAS.Business.Entity.scorecard, _
            ByVal parMode As FormMode, _
            ByVal parUserLogin As String _
    ) As Boolean
        With New DataAccess.scorecard()
            Try
                Return .InsertScorecard(parSCCode, parMode, parUserLogin)
            Catch ex As Exception
                Throw ex
            End Try
        End With
    End Function

    Function CekScorecardExsist(ByVal parSCCode As String) As Boolean
        With New DataAccess.scorecard()
            Try
                Return .CekScorecardExsist(parSCCode)
            Catch ex As Exception
                Throw ex
            End Try
        End With
    End Function

    Function ValidateRule(ByVal parRule As String, ByRef parMsg As String) As Boolean
        With New DataAccess.scorecard()
            Try
                Return .ValidateRule(parRule)
            Catch ex As Exception
                parMsg = ex.Message.Replace("'", "\'").Replace(vbCrLf, "\n")
                Return False
            End Try
        End With
    End Function

    Function DeleteSCH(ByVal parSCHID As String, _
            ByVal parMode As FormMode, _
            ByVal parUserLogin As String _
    ) As Boolean
        With New DataAccess.scorecard()
            Try
                Return .DeleteSCH(parSCHID, parMode, parUserLogin)
            Catch ex As Exception
                Throw ex
            End Try
        End With
    End Function

    Function DeleteScorecard(ByVal parSCCode As String, _
            ByVal parMode As FormMode, _
            ByVal parUserLogin As String _
    ) As Boolean
        With New DataAccess.scorecard()
            Try
                Return .DeleteScorecard(parSCCode, parMode, parUserLogin)
            Catch ex As Exception
                Throw ex
            End Try
        End With
    End Function


    Public Function GetSummaryField() As DataTable
        With New AAS.DataAccess.scorecard()
            Return .GetSummaryField()
        End With
    End Function

    Public Function GetCHSummaryField() As DataTable
        With New AAS.DataAccess.scorecard()
            Return .GetCHSummaryField()
        End With
    End Function

    Shared Function isDBTypeNumeric(ByVal pDBType As Int16) As Boolean
        Select Case pDBType
            Case 0, 1, 2
                Return False
            Case Else
                Return True
        End Select
    End Function
End Class
