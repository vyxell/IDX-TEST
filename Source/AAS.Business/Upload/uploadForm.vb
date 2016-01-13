Public Class uploadForm
    Function GetListOfFile() As DataTable
        With New DataAccess.uploadForm()
            Return .GetListOfFile()
        End With
    End Function

    Function GetLastMFormDate() As DataTable
        With New DataAccess.uploadForm()
            Return .GetLastMFormDate()
        End With
    End Function

    Function insertMForm(ByVal parData As AAS.Business.Entity.form_upload, ByVal parLogin As String)
        With New DataAccess.uploadForm
            Try
                Return .insertMForm(parData, parLogin)
            Catch ex As Exception
                Throw ex
            End Try
        End With
    End Function
End Class
