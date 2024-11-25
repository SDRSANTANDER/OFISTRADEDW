Public Class EntDocumentoExtendidoLin

    'Cabecera
    Public Property ObjectType As String = ""
    Public Property DocEntry As Integer
    Public Property DocNum As Integer
    Public Property NumAtCard As String = ""

    'Líneas
    Public Property LineNum As Integer
    Public Property BaseType As Integer
    Public Property BaseEntry As Integer
    Public Property BaseLine As Integer
    Public Property LineStatus As String = ""
    Public Property Articulo As String = ""
    Public Property Descripcion As String = ""
    Public Property RefExt As String = ""
    Public Property CuentaContable As String = ""
    Public Property Cantidad As Double
    Public Property CantidadPte As Double
    Public Property Precio As Double
    Public Property PorcDto As Double
    Public Property PrecioFinal As Double
    Public Property LineTotal As Double
    Public Property LineTotalFC As Double
    Public Property LineTotalPte As Double
    Public Property LineTotalPteFC As Double
    Public Property VatGroup As String = ""
    Public Property VatPorc As Double
    Public Property Moneda As String = ""
    Public Property Rate As Double
    Public Property TaxOnly As String = ""
    Public Property TaxCode As String = ""
    Public Property WTLiable As String = ""
    Public Property Almacen As String = ""
    Public Property Proyecto As String = ""
    Public Property Centro1Coste As String = ""
    Public Property Centro2Coste As String = ""
    Public Property Centro3Coste As String = ""
    Public Property Centro4Coste As String = ""
    Public Property Centro5Coste As String = ""

End Class
