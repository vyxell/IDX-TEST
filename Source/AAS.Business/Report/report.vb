Imports CrystalDecisions.Shared
Public Class report
    Function GetServerName() As String
        With New Business.Entity.report()
            Return .ServerName
        End With
    End Function

    Function GetListOfProduct() As DataTable
        With New DataAccess.report()
            Return .GetListOfProduct()
        End With
    End Function

    Function GetCARDataSet(ByVal parScorecard As String, ByVal parMonthFrom As String, ByVal parYearFrom As String, ByVal parMonthTo As String, ByVal parYearTo As String) As DataSet
        With New DataAccess.report()
            Return .GetCARDataSet(parScorecard, parMonthFrom, parYearFrom, parMonthTo, parYearTo)
        End With
    End Function

    Function GetGBKSDataSet(ByVal parScorecard As String, ByVal parMonthFrom As String, ByVal parYearFrom As String, ByVal parMonthTo As String, ByVal parYearTo As String) As DataSet
        With New DataAccess.report()
            Return .GetGBKSDataSet(parScorecard, parMonthFrom, parYearFrom, parMonthTo, parYearTo)
        End With
    End Function

    Function GetPopStabilityDataSet(ByVal parScorecard As String, ByVal parMonthFrom As String, ByVal parYearFrom As String, ByVal parMonthTo As String, ByVal parYearTo As String) As DataSet
        With New DataAccess.report()
            Return .GetPopStabilityDataSet(parScorecard, parMonthFrom, parYearFrom, parMonthTo, parYearTo)
        End With
    End Function

    Function GetOATRDataSet(ByVal parProduct As String, ByVal parMonthFrom As String, ByVal parYearFrom As String, ByVal parMonthTo As String, ByVal parYearTo As String) As DataSet
        With New DataAccess.report()
            Return .GetOATRDataSet(parProduct, parMonthFrom, parYearFrom, parMonthTo, parYearTo)
        End With
    End Function

    Function GetRRPDataSet(ByVal parType As String, ByVal parProduct As String, ByVal parMonthFrom As String, ByVal parYearFrom As String, ByVal parMonthTo As String, ByVal parYearTo As String) As DataSet
        With New DataAccess.report()
            Return .GetRRPDataSet(parType, parProduct, parMonthFrom, parYearFrom, parMonthTo, parYearTo)
        End With
    End Function

    Function GetRRPDailyDataSet(ByVal parProduct As String, ByVal parDateFrom As String, ByVal parDateTo As String) As DataSet
        With New DataAccess.report()
            Return .GetRRPDailyDataSet(parProduct, parDateFrom, parDateTo)
        End With
    End Function

    Function GetAWRDataSet(ByVal parDateFrom As String, ByVal parDateTo As String) As DataSet
        With New DataAccess.report()
            Return .GetAWRDataSet(parDateFrom, parDateTo)
        End With
    End Function

    Function GetRRCTDataSet(ByVal parType As String, ByVal parProduct As String, ByVal parMonthFrom As String, ByVal parYearFrom As String, ByVal parMonthTo As String, ByVal parYearTo As String) As DataSet
        With New DataAccess.report()
            Return .GetRRCTDataSet(parType, parProduct, parMonthFrom, parYearFrom, parMonthTo, parYearTo)
        End With
    End Function

    Function GetRRCTDailyDataSet(ByVal parProduct As String, ByVal parDateFrom As String, ByVal parDateTo As String) As DataSet
        With New DataAccess.report()
            Return .GetRRCTDailyDataSet(parProduct, parDateFrom, parDateTo)
        End With
    End Function

    Function GetRRORDataSet(ByVal parProduct As String, ByVal parMonth As String, ByVal parYear As String) As DataSet
        With New DataAccess.report()
            Return .GetRRORDataSet(parProduct, parMonth, parYear)
        End With
    End Function

    Function GetRRADataset(ByVal parProduct As String, ByVal parDateFrom As String, ByVal parDateTo As String) As DataSet
        With New DataAccess.report()
            Return .GetRRADataSet(parProduct, parDateFrom, parDateTo)
        End With
    End Function

    Function GetRRDDailyDataset(ByVal parProduct As String, ByVal parDateFrom As String, ByVal parDateTo As String) As DataSet
        With New DataAccess.report()
            Return .GetRRDDailyDataSet(parProduct, parDateFrom, parDateTo)
        End With
    End Function

    Function GetRRDDataset(ByVal parType As String, ByVal parProduct As String, ByVal parMonthFrom As String, ByVal parYearFrom As String, ByVal parMonthTo As String, ByVal parYearTo As String) As DataSet
        With New DataAccess.report()
            Return .GetRRDDataSet(parType, parProduct, parMonthFrom, parYearFrom, parMonthTo, parYearTo)
        End With
    End Function

    'Shared Function SetReportParam(ByVal pName As String, ByVal pValue As Object) As ParameterField
    '    Dim lDescValue As ParameterDiscreteValue
    '    Dim lParamField As ParameterField

    '    lDescValue = New ParameterDiscreteValue()
    '    lDescValue.Value = pValue
    '    lParamField = New ParameterField()
    '    lParamField.CurrentValues.Add(lDescValue)
    '    lParamField.ParameterFieldName = pName

    '    Return lParamField
    'End Function

    'Shared Function SetReportParamValues(ByVal parValue As Object) As ParameterValues
    '    Dim lDescValue As ParameterDiscreteValue
    '    Dim lParamValues As ParameterValues

    '    lDescValue = New ParameterDiscreteValue()
    '    lDescValue.Value = parValue
    '    lParamValues = New ParameterValues()
    '    lParamValues.Add(lDescValue)

    '    Return lParamValues
    'End Function

    Shared Function FilterNumericNull(ByVal pNum As Object) As Double
        If IsDBNull(pNum) Then
            Return 0
        Else
            Return pNum
        End If

    End Function
End Class
