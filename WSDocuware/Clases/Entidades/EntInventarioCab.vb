Public Class EntInventarioCab

    'DOCUWARE
    Public Property UserDW As String = ""           'Usuario Docuware
    Public Property PassDW As String = ""           'Password Docuware

    Public Property DOCIDDW As String = ""          'ID Docuware
    Public Property DOCURLDW As String = ""         'URL Docuware


    'CABECERA
    Public Property NIFEmpresa As String = ""       'Identificador fiscal de la empresa
    Public Property Opcion As String = ""           'C->Crea documento, U->Actualiza datos docuware
    Public Property Ambito As String = ""           'C->Compras, V->Ventas, I->Inventario
    Public Property NIFTercero As String = ""       'Identificador fiscal de tercero 
    Public Property RazonSocial As String = ""      'Razón social proveedor (CardName)
    Public Property DocNum As String = ""           'Número de documento emitido
    Public Property DocDate As Integer              'Formato YYYYMMDD (Es la fecha de documento, se mapea conta TaxDate)
    Public Property AccountDate As Integer          'Formato YYYYMMDD (Es la fecha contable, se mapea contra DocDate)

    Public Property Ref2 As String = ""             'Referencia cliente
    Public Property Comments As String = ""         'Comentarios

    Public Property JournalMemo As String = ""      'Entrada diario


    'DOCUMENTO DESTINO
    Public Property ObjTypeDestino As Integer       'Tipo de objeto destino según nomenclatura estándar SAP (13->Facturas, 15->Entregas, 22->Pedidos...)
    Public Property Draft As String = ""            'S->Borrador, N->Documento definitivo, I->Configuración interlocutor

    'DOCUMENTO ORIGEN
    Public Property RefOrigen As String = ""        'Número de albaranes/pedidos origen
    Public Property TipoRefOrigen As String = ""    '"DocNum","DocEntry" o "NumAtCard"

    'ALMACENES
    Public Property AlmacenOrigen As String = ""    'Almacén origen
    Public Property AlmacenDestino As String = ""   'Almacén destino

    'SAP
    Public Property UserSAP As String = ""
    Public Property PassSAP As String = ""

    'Líneas
    Public Property Lineas As New List(Of EntInventarioLin)

End Class
