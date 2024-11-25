Public Class EntLoteSAP

    'DOCUWARE
    Public Property UserDW As String = ""               'Usuario Docuware
    Public Property PassDW As String = ""               'Password Docuware

    'Generales
    Public Property NIFEmpresa As String = ""

    Public Property Ambito As String = ""               'C->Compras, V->Ventas, I->Inventario
    Public Property NIFTercero As String = ""           'Identificador fiscal de tercero 
    Public Property RazonSocial As String = ""          'Razón social proveedor (CardName)

    Public Property ObjTypeOrigen As Integer            'Tipo de objeto origen según nomenclatura estándar SAP (0->Por defecto, 13->Facturas, 15->Entregas, 22->Pedidos...)
    Public Property TipoRefOrigen As String = ""        'Tipo de referencia origen: "DocNum","DocEntry" o "NumAtCard"
    Public Property RefOrigen As String = ""            'Valor de la referencia del documento origen

    Public Property Articulo As String = ""             'Código de artículo
    Public Property RefExt As String = ""               'Referencia proveedor
    Public Property LineNum As Integer                  'Número de línea

    Public Property NumLote As String = ""

    Public Property CamposUsuario As New List(Of EntCampoUsuario)

    'SAP
    Public Property UserSAP As String = ""
    Public Property PassSAP As String = ""

End Class
