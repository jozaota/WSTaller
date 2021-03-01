Imports System.Data.Odbc
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Xml
Imports SAPbobsCOM

Module MdlSql

    Public SQLConexion As OdbcConnection
    Private cadenaOdbc As String
    Public errnum As Integer
    Public errdesc As String
    Public XmlDoc As XmlDocument
    Public ObSbo As SAPbobsCOM.Company

    Public Structure Credenciales

        Shared SqlServidor As String
        Shared SqlBaseDeDatos As String
        Shared SqlUsuario As String
        Shared SqlPassword As String
        Shared SapUsuario As String
        Shared SapPassword As String
        Shared SapServidorDeLicencia As String
        Shared SqlTipoConexion As String
        Shared HanaConexion As String
        Shared ClienteSql As String
        Shared SuperUsuario As Boolean
        Shared SqlConexion As Odbc.OdbcConnection
        Shared Motor As String
        Shared SerieOv As String
        Shared CtaPago As String
        Shared CondPago As String

    End Structure

    Public Function ConectaSql() As Boolean

        Try

            ConectaODBC()

            If Credenciales.Motor = "9" Then

                If (IntPtr.Size = 8) Then
                    Credenciales.HanaConexion = String.Concat(Credenciales.HanaConexion, "Driver={HDBODBC};")
                Else
                    Credenciales.HanaConexion = String.Concat(Credenciales.HanaConexion, "Driver={HDBODBC32};")
                End If

                Credenciales.HanaConexion = String.Concat(Credenciales.HanaConexion, "ServerNode=", Credenciales.SqlServidor & ";")
                Credenciales.HanaConexion = String.Concat(Credenciales.HanaConexion, "UID=", Credenciales.SqlUsuario, ";")
                Credenciales.HanaConexion = String.Concat(Credenciales.HanaConexion, "PWD=", Credenciales.SqlPassword, ";")

                Credenciales.SqlConexion = New OdbcConnection(Credenciales.HanaConexion)

                If Credenciales.SqlConexion.State = ConnectionState.Closed Then
                    Credenciales.SqlConexion.Open()
                End If
                If Credenciales.SqlConexion.State = ConnectionState.Open Then
                    Credenciales.SqlConexion.Close()
                End If
                Return True

            End If

            If Credenciales.Motor <> "9" Then

                Try

                    Credenciales.SqlConexion = New OdbcConnection(Credenciales.ClienteSql & "; Server= " & Credenciales.SqlServidor & "; Database=" & Credenciales.SqlBaseDeDatos & "; Uid=" & Credenciales.SqlUsuario & "; Pwd=" & Credenciales.SqlPassword)
                    If Credenciales.SqlConexion.State = ConnectionState.Closed Then
                        Credenciales.SqlConexion.Open()
                    End If
                    If Credenciales.SqlConexion.State = ConnectionState.Open Then
                        Credenciales.SqlConexion.Close()
                    End If
                    Return True

                Catch ex As Exception
                    'Credenciales.SqlTipoConexion = TipoConexion.Hana : Credenciales.SqlUsuario = "SYSTEM" : Credenciales.SqlPassword = "Passw0rd"
                End Try

            End If


            Return False

        Catch ex As Exception

            ' Errores.Mensaje = "Conexión SQL SERVER: " & ex.Message
            Return False

        End Try

    End Function

    Public Function DesConectaSQL(Optional ByRef mensaje As String = "") As Boolean

        Try

            If Credenciales.SqlConexion.State = ConnectionState.Open Then
                Credenciales.SqlConexion.Close()
            End If
            Return True

        Catch ex As Exception
            'mensaje = "DESSQLConexion SQL SERVER: " & ex.Message
            Return False
        End Try

    End Function

    Public Function ObtenerColeccion(ByVal consulta As String, Optional ByVal keepOpen As Boolean = False,
                                 Optional ByRef mensaje As String = "") As DataTable

        Try

            Dim dtt As New DataTable

            If Credenciales.SqlConexion.State = ConnectionState.Closed Then
                Credenciales.SqlConexion.Open()
            End If

            Dim dapTable As New OdbcDataAdapter(consulta, Credenciales.SqlConexion)
            dapTable.SelectCommand.CommandTimeout = 0
            dapTable.Fill(dtt)

            If Not keepOpen Then
                If Credenciales.SqlConexion.State = ConnectionState.Open Then
                    Credenciales.SqlConexion.Close()
                End If
            End If
            Return dtt

        Catch ex As OdbcException

            REM Operación no exitosa
            'mensaje = "SQL SERVER: (" & ex.ErrorCode & ") " & ex.Message

            Return Nothing

        End Try

    End Function

    Public Function ObtenerValor(ByVal consulta As String, Optional ByVal keepOpen As Boolean = False,
                              Optional ByRef mensaje As String = "") As String

        Try

            If Credenciales.SqlConexion.State = ConnectionState.Closed Then
                Credenciales.SqlConexion.Open()
            End If

            Dim comando As New OdbcCommand(consulta, Credenciales.SqlConexion)
            comando.CommandTimeout = 0
            comando.CommandText = consulta

            Dim valor As String = comando.ExecuteScalar
            If valor Is Nothing Then valor = ""

            If Not keepOpen Then
                If Credenciales.SqlConexion.State = ConnectionState.Open Then
                    Credenciales.SqlConexion.Close()
                End If
            End If

            REM Retornar el valor.
            Return valor

        Catch ex As OdbcException

            REM Operación no exitosa.
            'mensaje = "SQL SERVER: (" & ex.ErrorCode & ") " & ex.Message

            Return ""

        End Try

    End Function

    Public Function LiberarObjeto(ByVal oObject As Object)

        Try

            If oObject IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oObject)
            End If

            oObject = Nothing
            GC.Collect()

            Return True
        Catch ex As Exception

            oObject = Nothing
            GC.Collect()

            Return True
        End Try

    End Function

    REM Esta función acomoda el formato de una fecha entregada.
    Public Function FormatearFecha(ByVal fecha As String, Optional ByVal fechaSimple As Boolean = False,
                                   Optional ByVal horaCero As Boolean = False, Optional ByVal tiempo As String = "") As String

        Dim pais As String = System.Globalization.RegionInfo.CurrentRegion.DisplayName
        Dim aa As System.Globalization.DateTimeFormatInfo = System.Globalization.DateTimeFormatInfo.CurrentInfo
        Dim formatofecha As String
        Dim separador As String
        Dim dia, mes, anno, fechasal As String

        formatofecha = aa.ShortDatePattern
        separador = aa.DateSeparator

        Dim HoraActual As String = Format(Hour(FormatDateTime(Now, DateFormat.LongTime)), "00") & ":" &
                                   Format(Minute(FormatDateTime(Now, DateFormat.LongTime)), "00") & ":" &
                                   Format(Second(FormatDateTime(Now, DateFormat.LongTime)), "00")

        If fecha.ToString.Split("/").Length = 1 And fecha.ToString.Split("-").Length = 1 Then

            anno = Strings.Left(fecha, 4)
            mes = Strings.Mid(fecha, 5, 2)
            dia = Strings.Right(fecha, 2)

            If fechaSimple Then
                fechasal = anno & "-" & mes & "-" & dia
            Else

                If Not horaCero Then
                    fechasal = "{ts'" & anno & "-" & mes & "-" & dia & " " & IIf(tiempo = "", HoraActual, tiempo) & "'}"
                Else
                    fechasal = "{ts'" & anno & "-" & mes & "-" & dia & " 00:00:00'}"
                End If

            End If

            Return fechasal

        End If

        Dim tipofecha As Date = CType(fecha, Date)

        anno = Strings.Right("00" & CType(tipofecha.Year, String), 4)
        mes = Strings.Right("00" & CType(tipofecha.Month, String), 2)
        dia = Strings.Right("00" & CType(tipofecha.Day, String), 2)

        If formatofecha = "dd-MM-yyyy" And formatofecha = "dd/MM/yyyy" And formatofecha <> "M/d/yyyy" Then

            'MessageBox.Show(formatofecha & " : " & "Formato de Fecha no válido" & vbCrLf & vbCrLf & "Formatos válidos: " & vbCrLf & vbCrLf & "   - dd-MM-yyyy" & vbCrLf & "   - dd/MM/yyyy" & vbCrLf & "   - M/d/yyyy", "Verifique configuración regional", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return ""
            Exit Function

        End If

        If fechaSimple Then

            fechasal = anno & "-" & mes & "-" & dia

        Else

            If Not horaCero Then

                fechasal = "{ts'" & anno & "-" & mes & "-" & dia & " " & IIf(tiempo = "", HoraActual, tiempo) & "'}"

            Else

                fechasal = "{ts'" & anno & "-" & mes & "-" & dia & " 00:00:00'}"

            End If

        End If

        Return fechasal

    End Function

    Public Sub ConectaODBC()

        Dim appSettings = ConfigurationManager.AppSettings

        If appSettings.Count > 0 Then

            For Each key In appSettings.AllKeys

                If key = "server" Then
                    Credenciales.SqlServidor = appSettings.Get(key)
                ElseIf key = "dbname" Then
                    Credenciales.SqlBaseDeDatos = appSettings.Get(key)
                ElseIf key = "usu" Then
                    Credenciales.SapUsuario = appSettings.Get(key)
                ElseIf key = "pass" Then
                    Credenciales.SapPassword = appSettings.Get(key)
                ElseIf key = "lincense" Then
                    Credenciales.SapServidorDeLicencia = appSettings.Get(key)
                ElseIf key = "dbusu" Then
                    Credenciales.SqlUsuario = appSettings.Get(key)
                ElseIf key = "dbpass" Then
                    Credenciales.SqlPassword = appSettings.Get(key)
                ElseIf key = "driver" Then
                    Credenciales.ClienteSql = appSettings.Get(key)
                ElseIf key = "motor" Then
                    Credenciales.Motor = appSettings.Get(key)
                End If

            Next

        End If

        'ConectaSql()

    End Sub

End Module
