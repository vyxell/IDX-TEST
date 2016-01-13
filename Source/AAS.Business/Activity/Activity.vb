Imports AAS.DataAccess

Public Class Activity : Inherits BusinessBase
    Function GetSearchAccount(ByVal parActivity As AAS.Business.Entity.Activity) As DataTable
        With New DataAccess.Activity()
            Return .GetSearchAccount(parActivity)
        End With
    End Function

    Function GetActivityDetail(ByVal parActivityId As Int64) As DataTable
        With New DataAccess.Activity()
            Return .GetActivityDetail(parActivityId)
        End With
    End Function

    Function GetCCDetail(ByVal parAccNo As Int64, ByVal parCCNo As Int64) As DataTable
        With New DataAccess.Activity()
            Return .GetCCDetail(parAccNo, parCCNo)
        End With
    End Function

    Function GetActivityDetailView(ByVal parActivityId As Int64) As DataTable
        With New DataAccess.Activity()
            Return .GetActivityDetailView(parActivityId)
        End With
    End Function

    Function GetActivityList(ByVal parGrpAcc As Int64) As DataTable
        With New DataAccess.Activity()
            Return .GetActivityList(parGrpAcc)
        End With
    End Function

    Function GetActivityDetail_FinalOffer(ByVal parActivityId As Int64) As DataTable
        With New DataAccess.Activity()
            Return .GetActivityDetail_FinalOffer(parActivityId)
        End With
    End Function

    Function GetActivityDetail_RecOffer(ByVal parActivityId As Int64) As DataTable
        With New DataAccess.Activity()
            Return .GetActivityDetail_RecOffer(parActivityId)
        End With
    End Function

    Function InsertActivity(ByVal parActivity As AAS.Business.Entity.Activity, _
            ByVal parFinalOffer As ArrayList, _
            ByVal parRecOffer As ArrayList, _
            ByVal parAccNo As Int64, _
            ByVal parCardNo As Int64, _
            ByVal parMode As FormMode, _
            ByVal parUserLogin As String, _
            ByRef parErrMessage As String, _
            ByRef parActivityId As Int64, _
            Optional ByVal parAType As String = "ExsistCard" _
    ) As Boolean
        Try
            With New DataAccess.Activity()
                Return .InsertActivitySearch(parActivity, parFinalOffer, parRecOffer, parAccNo, parCardNo, parMode, parUserLogin, parErrMessage, parActivityId, parAType)
            End With
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Function InsertActivity_Waiving(ByVal parActivity As AAS.Business.Entity.Activity, _
            ByVal parMode As FormMode, _
            ByVal parUserLogin As String, _
            ByRef parErrMessage As String _
    ) As Boolean
        With New DataAccess.Activity()
            Return .InsertActivity_Waiving(parActivity, parMode, parUserLogin, parErrMessage)
        End With
    End Function

    Function CekCCExsist(ByVal parCardNo As String) As Boolean
        With New DataAccess.Activity()
            Try
                Return .CekCCExsist(parCardNo)
            Catch ex As Exception
                Throw ex
            End Try
        End With
    End Function

End Class
