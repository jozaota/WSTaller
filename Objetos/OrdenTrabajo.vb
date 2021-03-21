Public Class OrdenTrabajo

    Public Property CodCli As String
    Public Property NomCli As String
    Public Property CodVeh As String
    Public Property NomVeh As String
    Public Property Marca As String
    Public Property TipVeh As String
    Public Property ClaVeh As String
    Public Property NumMot As String
    Public Property Modelo As String
    Public Property EstVeh As String
    Public Property Chapa As String
    Public Property Chasis As String
    Public Property RucCli As String
    Public Property DirCli As String
    Public Property Te1Cli As String
    Public Property Te2Cli As String
    Public Property CiuCli As String
    Public Property Mail As String
    Public Property Kms As Decimal
    Public Property Cono As String
    Public Property NivCom As String
    Public Property FopCon As String
    Public Property FopCre As String
    Public Property NroOt As String
    'Public Property FecDoc As String
    Public Property EstaOt As String
    Public Property EstRec As String
    Public Property EtapOt As String
    Public Property PrioOt As String
    'Public Property FecRec As String
    'Public Property FecCie As String
    'Public Property FecEnt As String
    'Public Property FeEsCu As String
    Public Property HorRec As String
    Public Property HorEnt As String
    Public Property HoEsCu As String
    Public Property TotHor As String
    Public Property TotCos As Decimal
    Public Property Alinea As String
    Public Property MantKm As Decimal
    Public Property TasRue As String
    Public Property Extin As String
    Public Property Baliza As String
    Public Property Herram As String
    Public Property LlaFue As String
    Public Property Gato As String
    Public Property Auxil As String
    Public Property ComCor As String
    'Public Property OTPer As List(Of OtPer)
    'Public Property OTRep As List(Of OtRep)
    'Public Property OTIma As List(Of OtIma)
    'Public Property OTMat As List(Of OtMat)
    'Public Property OTTec As List(Of OtTec)
    'Public Property OTCot As List(Of OtCot)
    'Public Property OTHis As List(Of OtHis)
    'Public Property OTTra As List(Of OtTra)

End Class

Public Class OtPer
    Public Property DesPer As String
    Public Property EstPer As String
    Public Property CanPer As Decimal
    Public Property Comper As String

End Class

Public Class OtRep
    Public Property TarRep As String
    Public Property DesRep As String
    Public Property HorRep As Decimal
    Public Property CosRep As Decimal
    Public Property Comen As String

End Class

Public Class OtIma
    Public Property RutIma As String
    Public Property NomIma As String

End Class

Public Class OtMat
    Public Property CodMat As String
    Public Property NomMat As String
    Public Property CanMat As Decimal
    Public Property EstMat As String

End Class

Public Class OtTec
    Public Property CodTec As String
    Public Property NomTec As String
    Public Property EstTra As String
    Public Property HorIni As String
    Public Property HorFin As String
    Public Property HorTot As Decimal
    Public Property CodDoc As Integer

End Class

Public Class OtCot
    Public Property CodDoc As Integer
    Public Property NroDoc As Integer
    Public Property CodCli As String
    Public Property NomCli As String
    Public Property FecDoc As String
    Public Property ComDoc As String
    Public Property MonDoc As String
    Public Property TotDoc As Decimal
    Public Property EstDoc As String

End Class

Public Class OtHis
    Public Property NroOt As String
    Public Property FecCot As String
    Public Property Comen As String

End Class

Public Class OtTra
    Public Property CodDoc As Integer
    Public Property NroDoc As Integer
    Public Property TipDoc As String
    Public Property ComDoc As String
    Public Property CodCli As String
    Public Property NomCli As String
    Public Property MonDoc As String
    Public Property TotDoc As Decimal

End Class

