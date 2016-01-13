
Imports AAS.DataAccess

Public Class Product : Inherits BusinessBase

    Function GetSelectionOfProduct(Optional ByVal parProductID As String = "") As DataTable
        With New DataAccess.Product()
            Return .GetSelectionOfProduct(parProductID)
        End With
    End Function

    Function GetDescriptionOfProduct(Optional ByVal parProductID As String = "") As DataTable
        With New DataAccess.Product
            Return .GetDescriptionOfProduct(parProductID)
        End With
    End Function

End Class
