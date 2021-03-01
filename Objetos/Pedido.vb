Public Class Pedido

    Public Property CardCode As String
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
    Public Property Detalles As List(Of PedidoDet)

End Class
