Imports AAS.DataAccess

Public Class User : Inherits BusinessBase
    Function GetListOfUser(Optional ByVal paramUserID As String = "") As DataTable
        With New DataAccess.User()
            Return .GetListOfUser(paramUserID)
        End With
    End Function

    Function CekPassword(ByVal paramUserID As String, ByVal paramPwd As String, ByVal paramDBPwd As String) As Boolean
        With New DataAccess.User()
            Return .CekPassword(paramUserID, paramPwd, paramDBPwd)
        End With
    End Function

    Function DeleteUser(ByVal paramUserID As String, _
            ByVal paramMode As FormMode, _
            ByVal paramUserLogin As String _
    ) As Boolean
        With New DataAccess.User()
            Try
                .DeleteUser(paramUserID, paramMode, paramUserLogin)
            Catch ex As Exception
                Throw ex
            End Try
        End With
    End Function

    Function InsertUser(ByVal paramUserID As String, _
            ByVal paramUserName As String, _
            ByVal paramUserNIP As String, _
            ByVal paramUserPassword As String, _
            ByVal paramAccessLevelID As String, _
            ByVal paramUserStatus As Boolean, _
            ByVal paramMode As FormMode, _
            ByVal paramUserLogin As String _
    ) As Boolean
        With New DataAccess.User()
            Return .InsertUser(paramUserID, paramUserName, paramUserNIP, paramUserPassword, paramAccessLevelID, paramUserStatus, paramMode, paramUserLogin)
        End With
    End Function

    Function InsertUserReset(ByVal paramUserID As String, _
        ByVal paramUserName As String, _
        ByVal paramUserNIP As String, _
        ByVal paramUserPassword As String, _
        ByVal paramAccessLevelID As String, _
        ByVal paramUserStatus As Boolean, _
        ByVal paramMode As FormMode, _
        ByVal paramUserLogin As String _
) As Boolean
        With New DataAccess.User()
            Return .InsertUserReset(paramUserID, paramUserName, paramUserNIP, paramUserPassword, paramAccessLevelID, paramUserStatus, paramMode, paramUserLogin)
        End With
    End Function

    Function getUserAccessLevel(ByVal paramUserName As String) As DataTable
        With New DataAccess.User()
            Return .getUserAccessLevel(paramUserName)
        End With
    End Function

    Sub insertLoginLog(ByVal paramUserID As String, ByVal paramIP As String, ByVal parAction As String)
        With New DataAccess.User()
            .insertLoginLog(paramUserID, paramIP, parAction)
        End With

    End Sub

    Sub updateLastAccess(ByVal paramUserID As String)
        With New DataAccess.User()
            .updateLastAccess(paramUserID)
        End With

    End Sub

    Function ChangePassword(ByVal paramUserID As String, ByVal paramPwd As String, ByVal paramPwdNew As String) As Boolean
        With New DataAccess.User()
            Return .ChangePassword(paramUserID, paramPwd, paramPwdNew)
        End With
    End Function

    Function CekUserExsist(ByVal paramUserName As String) As Boolean
        With New DataAccess.User()
            Try
                Return .CekUserExsist(paramUserName)
            Catch ex As Exception
                Throw ex
            End Try
        End With
    End Function

    Function CekUserLogin(ByVal paramUserName As String) As Boolean
        With New DataAccess.User()
            Try
                Return .CekUserLogin(paramUserName)
            Catch ex As Exception
                Throw ex
            End Try
        End With
    End Function

    Function ValidateUser(ByVal paramUserName As String) As Integer
        If (paramUserName.Length < 6 Or paramUserName.Length > 15) Then
            Return 2
        ElseIf (CekUserExsist(paramUserName)) Then
            Return 3
        End If
    End Function

    Function ValidateUserEdit(ByVal paramUserName As String) As Integer
        If (CekUserLogin(paramUserName)) Then
            Return 1
        End If
    End Function

    Function ValidatePassword(ByVal paramUser As String, ByVal paramPassword As String) As Integer
        If (paramPassword.Length < 6 Or paramPassword.Length > 15) Then
            Return 2
        ElseIf (Not isContainValid(paramPassword)) Then
            Return 3
        ElseIf (isSamePosition(paramUser, paramPassword)) Then
            Return 4
        End If
    End Function

    Function isSamePosition(ByVal paramUser As String, ByVal paramPass As String) As Boolean
        Dim lData As DataTable
        Dim lPass(4) As String
        Dim lIndex As Integer
        Dim lArrIndex As Integer
        Dim lResult As Boolean = False
        With New DataAccess.User()

            lData = .GetListOfUser(paramUser)
            If lData.Rows.Count > 0 Then
                With New Meta.Component.General()
                    lPass(0) = .Decrypt(lData.Rows(0)("usr_password"))
                    lPass(1) = .Decrypt(lData.Rows(0)("usr_pass_history1"))
                    lPass(2) = .Decrypt(lData.Rows(0)("usr_pass_history2"))
                    lPass(3) = .Decrypt(lData.Rows(0)("usr_pass_history3"))
                    lPass(4) = .Decrypt(lData.Rows(0)("usr_pass_history4"))
                End With
                For lArrIndex = 0 To 4
                    lIndex = 0
                    While lIndex < lPass(lArrIndex).Length And lIndex < paramPass.Length
                        If lPass(lArrIndex).Substring(lIndex, 1) = paramPass.Substring(lIndex, 1) Then
                            lResult = True
                            Exit While
                        End If
                        lIndex = lIndex + 1
                    End While
                    If lResult Then
                        Exit For
                    End If
                Next
            Else
                Return False
            End If
        End With
        Return lResult
    End Function

    Function isContainValid(ByVal paramText As String) As Boolean
        Dim lIndex As Int16
        Dim lReturnNum As Boolean = False
        Dim lReturnAlp As Boolean = False
        For lIndex = 0 To 9
            If paramText.IndexOf(lIndex) > -1 Then
                lReturnNum = True
            ElseIf Not IsNumeric(paramText) Then
                lReturnAlp = True
            End If
        Next
        Return lReturnNum And lReturnAlp
    End Function
End Class
