Public Class ModCliente

    Public Property CardCode As String
    Public Property CardName As String
    Public Property CardFName As String
    Public Property AddID As String
    Public Property VatStatus As String
    Public Property LicTradNum As String
    Public Property Phone1 As String
    Public Property E_mail As String
    Public Property Direccion As List(Of ModClienteDet)

End Class

Public Class ModClienteDet
    Public Property Address As String
    Public Property AddressType As String
    Public Property Block As String
    Public Property Street As String
    Public Property City As String
    Public Property State As String
    Public Property Country As String
End Class