Public Class EntDocumentoLin


    'BASE IMPONIBLE (SERVICIOS)
    Public Property CuentaContable As String = ""   'Cuenta contable para documentos de tipo servicio
    Public Property Concepto As String = ""         'Descripción tanto en documentos de tipo servicio como artículos
    Public Property Proyecto As String = ""         'Código de proyecto
    Public Property LineTotal As Double             'Total de la línea, incluidos impuestos
    Public Property BaseTotal As Double             'Base imponible de la línea, no usar
    Public Property VATTotal As Double              'IVA de la línea
    Public Property VATGroup As String = ""         'Código del grupo de IVA
    Public Property VATPorc As Double               'Porcentaje de IVA
    Public Property Intracomunitario As String = "" 'S/N
    Public Property TaxOnly As String = ""          'S/N
    Public Property TaxCode As String = ""          'Código del grupo de impuesto
    Public Property WTLiable As String = ""         'S/N (sujeto a retención)

    Public Property Almacen As String = ""          'Almacén

    Public Property Centro1Coste As String = ""     'Centro de coste 1
    Public Property Centro2Coste As String = ""     'Centro de coste 2, no usar si no hace falta
    Public Property Centro3Coste As String = ""     'Centro de coste 3, no usar si no hace falta
    Public Property Centro4Coste As String = ""     'Centro de coste 4, no usar si no hace falta
    Public Property Centro5Coste As String = ""     'Centro de coste 5, no usar si no hace falta

    Public Property ShipDate As Integer             'yyyyMMdd
    Public Property RequiredDate As Integer         'yyyyMMdd

    'ARTÍCULOS
    Public Property LineNum As Integer              'Número de línea
    Public Property VisOrder As Integer             'Orden visualización de línea
    Public Property Articulo As String = ""         'Código de artículo
    Public Property Cantidad As Double              'Cantidad
    Public Property PrecioUnidad As Double          'Precio unitario del artículo
    Public Property PorcentajeDescuento As Double   'Porcentaje de descuento, no usar (pasar 0)
    Public Property RefExt As String = ""           'Referencia proveedor

    'CAMPOS AÑADIDOS
    'Public Property RefExt As String = ""              'Referencia proveedor
    Public Property TipoRefOrigen As String = ""        'Tipo de referencia origen: "DocNum","DocEntry" o "NumAtCard"
    Public Property RefOrigen As String = ""            'Valor de la referencia del documento origen
    'Public Property NumOT As Integer
    'Public Property Operarios As String = ""
    Public Property CantidadOrigen As String = ""       'S/N
    'Public Property PrecioOrigen As String = ""        'S/N

    'Public Property LINURLDW As String = ""

    'CAMPOS DE USUARIO
    Public Property Lotes As New List(Of EntDocumentoLote)

    'Campos usuario
    Public Property CamposUsuario As New List(Of EntCampoUsuario)


End Class
