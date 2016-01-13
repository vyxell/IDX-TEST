Public Class upload : Inherits BusinessBase
    Function GetListOfUpload(Optional ByVal parUploadID As String = "") As DataTable
        With New DataAccess.upload()
            Return .GetListOfUpload(parUploadID)
        End With
    End Function

    Function InsertUpload(ByVal parUpload As AAS.Business.Entity.upload, _
            ByVal parMode As FormMode, _
            ByVal parUserLogin As String _
    ) As Boolean
        With New DataAccess.upload()
            Return .InsertUpload(parUpload, parMode, parUserLogin)
        End With
    End Function

    Function DeleteUpload(ByVal parUploadID As String, _
            ByVal parMode As FormMode, _
            ByVal parUserLogin As String _
    ) As Boolean
        With New DataAccess.upload()
            Return .DeleteUpload(parUploadID, parMode, parUserLogin)
        End With
    End Function

    Function DeleteFile(ByVal parPathFile As String) As Boolean
        Try
            With New DataAccess.upload()
                Return .DeleteFile(parPathFile)
            End With
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Function CekUploadExsist(Optional ByVal parFile As String = "") As Boolean
        With New DataAccess.upload()
            Return .CekUploadExsist(parFile)
        End With
    End Function

End Class
