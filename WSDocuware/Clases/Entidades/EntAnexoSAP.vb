Public Class EntAnexoSAP

    'DOCUWARE
    Public Property UserDW As String = ""               'Usuario Docuware
    Public Property PassDW As String = ""               'Password Docuware

    'CABECERA
    Public Property NIFEmpresa As String = ""           'Identificador fiscal de la empresa

    Public Property Ambito As String = ""               'C->Compras, V->Ventas, I->Inventario
    Public Property NIFTercero As String = ""           'Identificador fiscal de tercero 
    Public Property RazonSocial As String = ""          'Razón social proveedor (CardName)

    'DOCUMENTO DESTINO
    Public Property ObjType As Integer                  'Tipo de objeto destino según nomenclatura estándar SAP (13->Facturas, 15->Entregas, 22->Pedidos...)
    Public Property DOCIDDW As String = ""
    Public Property DocNum As String = ""
    Public Property NumAtCard As String = ""

    'ANEXO
    Public Property Nombre As String = ""               'Extensión incluida
    Public Property Base64 As String = ""

    'SAP
    Public Property UserSAP As String = ""
    Public Property PassSAP As String = ""

End Class
