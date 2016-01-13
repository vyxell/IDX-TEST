Imports AAS.DataAccess

Public Class Logs : Inherits BusinessBase
    Function GetListOfLogs(ByVal paramLogID As String, ByVal paramDateFrom As Date, ByVal paramDateTo As Date) As DataTable
        With New DataAccess.Logs()
            Return .GetListOfLogs(paramLogID, paramDateFrom, paramDateTo)
        End With
    End Function
End Class
