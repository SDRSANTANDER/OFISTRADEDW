Public Class EntAcuerdoLin

    'Artículo
    Public Property LineNum As Integer              'Número de línea
    Public Property VisOrder As Integer             'Orden visualización de línea
    Public Property Articulo As String = ""         'Código de artículo
    Public Property RefExt As String = ""           'Referencia proveedor
    Public Property Concepto As String = ""         'Descripción tanto en documentos de tipo servicio como artículos
    Public Property Cantidad As Double              'Cantidad
    Public Property PrecioUnidad As Double          'Precio unitario del artículo
    Public Property PorcentajeDescuento As Double   'Porcentaje de descuento, no usar (pasar 0)
    Public Property PorcentajeDevolucion As Double  'Porcentaje de devolución
    Public Property Moneda As String = ""           'Moneda

    'Monetarioa
    Public Property ImportePlanificado As Double    'Importe planificado

    'Genérico
    Public Property Proyecto As String = ""         'Código de proyecto
    Public Property Comentarios As String = ""      'Comentarios
    Public Property FechaGarantia As Double         'Fecha garantía yyyyMMdd


End Class
