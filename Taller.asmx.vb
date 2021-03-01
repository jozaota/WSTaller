Imports System.ComponentModel
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Configuration
Imports System.Security
Imports System.Xml
'Imports System.Web.Http
Imports SAPbobsCOM
Imports SBODI_Server
Imports WSTaller
Imports System.Web.Http

' Para permitir que se llame a este servicio web desde un script, usando ASP.NET AJAX, quite la marca de comentario de la línea siguiente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://tempuri.org/")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class Taller
    Inherits System.Web.Services.WebService

    Public Sub New()
        MyBase.New()

        'This call is required by the Web Services Designer.
        InitializeComponent()
        'Add your own initialization code after the InitializeComponent() call
        ConectaODBC()

    End Sub

    'Required by the Web Services Designer
    Private components As System.ComponentModel.IContainer
    Property Sbo As SAPbobsCOM.Company
    Property SboSrv As SAPbobsCOM.CompanyService
    Public oRecordSet As SAPbobsCOM.Recordset
    Public oRecordSetAux As SAPbobsCOM.Recordset
    Private oGeneralService As SAPbobsCOM.GeneralService
    Private oGeneralData As SAPbobsCOM.GeneralData
    Private oGeneralParams As SAPbobsCOM.GeneralDataParams
    Private oSons As SAPbobsCOM.GeneralDataCollection
    Private oSon As SAPbobsCOM.GeneralData

    'NOTE: The following procedure is required by the Web Services Designer
    'It can be modified using the Web Services Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        'CODEGEN: This procedure is required by the Web Services Designer
        'Do not modify it using the code editor.
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    Public Sub ConectaODBC()

        Dim appSettings = ConfigurationManager.AppSettings

        If appSettings.Count > 0 Then

            For Each key In appSettings.AllKeys

                If key = "server" Then
                    Credenciales.SqlServidor = appSettings.Get(key)
                ElseIf key = "dbname" Then
                    Credenciales.SqlBaseDeDatos = key
                ElseIf key = "usu" Then
                    Credenciales.SapUsuario = key
                ElseIf key = "pass" Then
                    Credenciales.SapPassword = key
                ElseIf key = "lincense" Then
                    Credenciales.SapServidorDeLicencia = key
                ElseIf key = "dbusu" Then
                    Credenciales.SqlUsuario = key
                ElseIf key = "dbpass" Then
                    Credenciales.SqlPassword = key
                ElseIf key = "driver" Then
                    Credenciales.ClienteSql = key
                ElseIf key = "motor" Then
                    Credenciales.Motor = key
                End If

            Next

        End If

        ConectaSql()

    End Sub


    <WebMethod()> _
    Public Function HelloWorld() As String
        Return "Hola a todos"
    End Function

    <WebMethod>
    Public Function AgregarOT(<FromBody> ByVal value As OrdenTrabajo) As Integer

        Dim OTREP As String = String.Empty
        Dim OTPER As String = String.Empty
        Dim OTIMA As String = String.Empty
        Dim OTMAT As String = String.Empty
        Dim OTTEC As String = String.Empty
        Dim OTCOT As String = String.Empty
        Dim OTHIS As String = String.Empty
        Dim OTTRA As String = String.Empty

        Try

            OTREP = ObtenerValor("SELECT SonName FROM UDO1 WHERE Code = 'OT' AND TableName = 'OTREP'")
            OTPER = ObtenerValor("SELECT SonName FROM UDO1 WHERE Code = 'OT' AND TableName = 'OTPER'")
            OTIMA = ObtenerValor("SELECT SonName FROM UDO1 WHERE Code = 'OT' AND TableName = 'OTIMA'")
            OTMAT = ObtenerValor("SELECT SonName FROM UDO1 WHERE Code = 'OT' AND TableName = 'OTMAT'")
            OTTEC = ObtenerValor("SELECT SonName FROM UDO1 WHERE Code = 'OT' AND TableName = 'OTTEC'")
            OTCOT = ObtenerValor("SELECT SonName FROM UDO1 WHERE Code = 'OT' AND TableName = 'OTCOT'")
            OTHIS = ObtenerValor("SELECT SonName FROM UDO1 WHERE Code = 'OT' AND TableName = 'OTHIS'")
            OTTRA = ObtenerValor("SELECT SonName FROM UDO1 WHERE Code = 'OT' AND TableName = 'OTTRA'")

            oGeneralService = SboSrv.GetGeneralService("OT")
            oGeneralData = oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralData)
            'oGeneralParams = oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralDataParams)
            'oGeneralParams.SetProperty("DocEntry", CodGot)
            'oGeneralData = oGeneralService.GetByParams(oGeneralParams)

            oSons = oGeneralData.Child(OTREP)



            Return 1
        Catch ex As Exception
            Return 0
        End Try

    End Function


End Class