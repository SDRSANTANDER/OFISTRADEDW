Public Class EntArticulo

    Public Property ID As String = ""
    Public Property Descripcion As String = ""
    Public Property Activo As String = ""
    Public Property Proveedor As String = ""
	Public Property ReferenciaExterna As String
    Public Property Tipo As String = ""
    Public Property Grupo As String = ""

    Public Property Venta As String = ""
    Public Property Compra As String = ""
    Public Property Inventario As String = ""

    Public Property UltimaCompraPrecio As Double
    Public Property UltimaCompraMoneda As String = ""

End Class
