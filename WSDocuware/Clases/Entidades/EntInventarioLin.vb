Public Class EntInventarioLin

    'ARTÍCULOS
    Public Property Articulo As String = ""         'Código de artículo
    Public Property RefExt As String = ""           'Referencia proveedor
    Public Property Concepto As String = ""         'Descripción tanto en documentos de tipo servicio como artículos
    Public Property Cantidad As Double              'Cantidad
    Public Property PrecioUnidad As Double          'Precio unitario del artículo
    Public Property Almacen As String = ""          'Almacén

    'PROYECTO
    Public Property Proyecto As String = ""         'Código de proyecto

    'CENTROS COSTE
    Public Property Centro1Coste As String = ""     'Centro de coste 1
    Public Property Centro2Coste As String = ""     'Centro de coste 2, no usar si no hace falta
    Public Property Centro3Coste As String = ""     'Centro de coste 3, no usar si no hace falta
    Public Property Centro4Coste As String = ""     'Centro de coste 4, no usar si no hace falta
    Public Property Centro5Coste As String = ""     'Centro de coste 5, no usar si no hace falta

    'CAMPOS DE USUARIO
    Public Property Lotes As New List(Of EntDocumentoLote)

End Class
