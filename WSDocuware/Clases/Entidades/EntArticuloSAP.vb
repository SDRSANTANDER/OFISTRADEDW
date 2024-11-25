Public Class EntArticuloSAP

    'DOCUWARE
    Public Property UserDW As String = ""               'Usuario Docuware
    Public Property PassDW As String = ""               'Password Docuware

    'Generales
    Public Property NIFEmpresa As String = ""
    Public Property Opcion As String = ""               'C->Crea artículo, U->Actualiza artículo

    'Campos SAP
    Public Property Serie As Integer
    Public Property Codigo As String = ""
    Public Property Nombre As String = ""
    Public Property NombreExtranjero As String = ""
    Public Property Tipo As String = ""                 'Se considera si es numérico
    Public Property Grupo As String = ""                'Se considera si es numérico
    Public Property GrupoUnidadMedida As String = ""    'Se considera si es numérico
    Public Property CodigoBarras As String = ""

    Public Property Venta As String = ""                'Y/N 
    Public Property Compra As String = ""               'Y/N
    Public Property Inventario As String = ""           'Y/N

    'Datos generales
    Public Property GeneralSujetoRetencion As String = ""               'Y/N            
    Public Property GeneralFabricante As String = ""                    'Se considera si es numérico
    Public Property GeneralClaseExpedicion As String = ""               'Se considera si es numérico
    Public Property GeneralGestionarArticuloPorLotes As String = ""     'Y/N 
    Public Property GeneralRelevanteInstratat As String = ""            'Y/N 
    Public Property GeneralActivo As String = ""                        'S/N

    'Datos de compra
    Public Property CompraProveedorDefecto As String = ""
    Public Property CompraNumeroCatalogoFabricante As String = ""
    Public Property CompraUnidadMedida As String = ""
    Public Property CompraArticulosUnidad As Double
    Public Property CompraUnidadMedidaEmbalaje As String = ""
    Public Property CompraCantidadEmbalaje As Double
    Public Property CompraGrupoImpositivo As String = ""

    'Datos de venta
    Public Property VentaUnidadMedida As String = ""
    Public Property VentaArticulosUnidad As Double
    Public Property VentaUnidadMedidaEmbalaje As String = ""
    Public Property VentaCantidadEmbalaje As Double
    Public Property VentaGrupoImpositivo As String = ""

    'Datos de inventario
    Public Property InventarioCuentasDeMayorPor As String = ""          'Se considera si es numérico
    Public Property InventarioUnidadMedida As String = ""
    Public Property InventarioPeso As Double
    Public Property InventarioGestionStockAlmacen As String = ""        'Y/N 

    'Datos de activo fijo
    Public Property ActivoFijoClase As String = ""
    Public Property ActivoFijoGrupo As String = ""
    Public Property ActivoFijoGrupoAmortizacion As String = ""
    Public Property ActivoFijoNumeroInventario As String = ""
    Public Property ActivoFijoNumeroSerie As String = ""
    Public Property ActivoFijoEmplazamiento As Integer
    Public Property ActivoFijoTecnico As Integer
    Public Property ActivoFijoEmpleado As Integer
    Public Property ActivoFijoFechaCapitalizacion As Integer    'Formato YYYYMMDD 
    Public Property ActivoFijoEstadistico As String = ""        'Y/N
    Public Property ActivoFijoCesion As String = ""             'Y/N
    Public Property ActivoFijoAreaValoracion As String = ""
    Public Property ActivoFijoFechaInicio As Integer            'Formato YYYYMMDD 
    Public Property ActivoFijoTipo As String = ""
    Public Property ActivoFijoAnyoFiscal As String = ""
    Public Property ActivoFijoVidaUtil As Integer
    Public Property ActivoFijoVidaUtilUnidades As Integer
    Public Property ActivoFijoCAPHistorico As Double

    'Comentarios
    Public Property Comentarios As String = ""

    'Campos de usuario
    Public Property CamposUsuario As New List(Of EntCampoUsuario)

    'SAP
    Public Property UserSAP As String = ""
    Public Property PassSAP As String = ""

End Class

