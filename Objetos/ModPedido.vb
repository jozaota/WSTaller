Public Class ModPedido

    Public Property DocNum As Integer
    Public Property Fecha As String
    Public Property FechaEntrega As String
    Public Property TrnspCode As String
    Public Property NumAtCard As String
    Public Property GroupNum As String
    Public Property Address As String
    Public Property Address2 As String
    Public Property Comments As String
    Public Property DocTotal As Double
    Public Property FechaCobro As String
    Public Property CounterRef As String
    Public Property Detalles As List(Of ModPedidoDet)


End Class

Public Class ModPedidoDet
    Public Property ItemCode As String
    Public Property Quantity As Double
    Public Property Price As Double
    Public Property TaxCode As String
    Public Property DiscPrcnt As Double
    Public Property LineTotal As Double
End Class
