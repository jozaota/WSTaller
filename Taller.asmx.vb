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

            oGeneralData.SetProperty("U_CODCLI", value.CodCli)
            oGeneralData.SetProperty("U_NOMCLI", value.NomCli)
            oGeneralData.SetProperty("U_CODVEH", value.CodVeh)
            oGeneralData.SetProperty("U_NOMVEH", value.NomVeh)
            oGeneralData.SetProperty("U_MARCA", value.Marca)
            oGeneralData.SetProperty("U_TIPVEH", value.TipVeh)
            oGeneralData.SetProperty("U_CLAVEH", value.ClaVeh)
            oGeneralData.SetProperty("U_NUMMOT", value.NumMot)
            oGeneralData.SetProperty("U_MODELO", value.Modelo)
            oGeneralData.SetProperty("U_ESTVEH", value.EstVeh)
            oGeneralData.SetProperty("U_CHAPA", value.Chapa)
            oGeneralData.SetProperty("U_CHASIS", value.Chasis)
            oGeneralData.SetProperty("U_RUCCLI", value.RucCli)
            oGeneralData.SetProperty("U_DIRCLI", value.DirCli)
            oGeneralData.SetProperty("U_TE1CLI", value.Te1Cli)
            oGeneralData.SetProperty("U_TE2CLI", value.Te2Cli)
            oGeneralData.SetProperty("U_CIUCLI", value.CiuCli)
            oGeneralData.SetProperty("U_MAIL", value.Mail)
            oGeneralData.SetProperty("U_KMS", value.Kms)
            oGeneralData.SetProperty("U_CONO", value.Cono)
            oGeneralData.SetProperty("U_NIVCOM", value.NivCom)
            oGeneralData.SetProperty("U_FOPCON", value.FopCon)
            oGeneralData.SetProperty("U_FOPCRE", value.FopCre)
            oGeneralData.SetProperty("U_NROOT", value.NroOt)
            oGeneralData.SetProperty("U_FECDOC", value.FecDoc)
            oGeneralData.SetProperty("U_ESTAOT", value.EstaOt)
            oGeneralData.SetProperty("U_ESTREC", value.EstRec)
            oGeneralData.SetProperty("U_ETAPOT", value.EtapOt)
            oGeneralData.SetProperty("U_PRIOOT", value.PrioOt)
            oGeneralData.SetProperty("U_FECREC", value.FecRec)
            oGeneralData.SetProperty("U_FECCIE", value.FecCie)
            oGeneralData.SetProperty("U_FECENT", value.FecEnt)
            oGeneralData.SetProperty("U_FEESCU", value.FeEsCu)
            oGeneralData.SetProperty("U_HORREC", value.HorRec)
            oGeneralData.SetProperty("U_HORENT", value.HorEnt)
            oGeneralData.SetProperty("U_HOESCU", value.HoEsCu)
            oGeneralData.SetProperty("U_TOTHOR", value.TotHor)
            oGeneralData.SetProperty("U_TOTCOS", value.TotCos)
            oGeneralData.SetProperty("U_ALINEA", value.Alinea)
            oGeneralData.SetProperty("U_MANTKM", value.MantKm)
            oGeneralData.SetProperty("U_TASRUE", value.TasRue)
            oGeneralData.SetProperty("U_EXTIN", value.Extin)
            oGeneralData.SetProperty("U_BALIZA", value.Baliza)
            oGeneralData.SetProperty("U_HERRAM", value.Herram)
            oGeneralData.SetProperty("U_LLAFUE", value.LlaFue)
            oGeneralData.SetProperty("U_GATO", value.Gato)
            oGeneralData.SetProperty("U_AUXIL", value.Auxil)
            oGeneralData.SetProperty("U_COMCOR", value.ComCor)

            If value.OTRep.Count > 0 Then
                oSons = oGeneralData.Child(OTREP)

                For i As Integer = 0 To value.OTRep.Count - 1
                    oSon = oSons.Add()
                    oSon.SetProperty("U_TARREP", value.OTRep(i).TarRep)
                    oSon.SetProperty("U_DESREP", value.OTRep(i).DesRep)
                    oSon.SetProperty("U_HORREP", value.OTRep(i).HorRep)
                    oSon.SetProperty("U_COSREP", value.OTRep(i).CosRep)
                    oSon.SetProperty("U_COMEN", value.OTRep(i).Comen)
                Next
            End If

            oGeneralService.Add(oGeneralData)

            LiberarObjeto(oGeneralService)


            Return 1
        Catch ex As Exception
            Return 0
        End Try

    End Function


End Class