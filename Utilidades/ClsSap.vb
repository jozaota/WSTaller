Imports SAPbobsCOM

Public Class ClsSap

    Public Shared ObSbo As SAPbobsCOM.Company
    Public Shared ObSboServ As SAPbobsCOM.CompanyService

    Public Shared Function ConectaDI() As Boolean
        Try
            ObSbo = New SAPbobsCOM.Company
            ObSbo.Server = Credenciales.SqlServidor
            ObSbo.CompanyDB = Credenciales.SqlBaseDeDatos
            ObSbo.DbServerType = Credenciales.Motor 'SAPbobsCOM.BoDataServerTypes.dst_MSSQL2016
            ObSbo.UserName = Credenciales.SapUsuario
            ObSbo.Password = Credenciales.SapPassword
            ObSbo.DbUserName = Credenciales.SqlUsuario
            ObSbo.DbPassword = Credenciales.SqlPassword
            ObSbo.language = BoSuppLangs.ln_Spanish_La

            ObSbo.UseTrusted = False
            errnum = ObSbo.Connect
            If errnum <> 0 Then
                ObSbo.GetLastError(errnum, errdesc)
                Throw New Exception(errdesc)
            End If

            ObSboServ = ObSbo.GetCompanyService

            Return True

        Catch ex As Exception
            'MessageBox.Show(sError, "Error")
            Return False
        End Try
    End Function

    Public Shared Function DesconectaDI() As Boolean
        Try

            ObSbo.Disconnect()

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

End Class
