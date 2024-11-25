Public Class EntAcuerdoCab

    'DOCUWARE
    Public Property UserDW As String = ""           'Usuario Docuware
    Public Property PassDW As String = ""           'Password Docuware

    Public Property DOCIDDW As String = ""          'ID Docuware
    Public Property DOCURLDW As String = ""         'URL Docuware

    'Generales
    Public Property NIFEmpresa As String = ""       'Identificador fiscal de la empresa
    Public Property Opcion As String = ""           'C->Crea acuerdo, U->Actualiza acuerdo

    'CABECERA
    Public Property ObjTypeDestino As Integer       'Tipo de objeto destino según nomenclatura estándar SAP (13->Facturas, 15->Entregas, 22->Pedidos...)
    Public Property Ambito As String = ""           'C->Compras, V->Ventas, I->Inventario
    Public Property NIFTercero As String = ""       'Identificador fiscal de tercero 
    Public Property RazonSocial As String = ""      'Razón social proveedor (CardName)
    Public Property PersonaContacto As Integer      'Persona contacto
    Public Property NumAtCard As String = ""        'Referencia interlocutor comercial
    Public Property Numero As String = ""           'Número acuerdo global
    Public Property Proyecto As String = ""         'Proyecto
    Public Property FechaInicio As Integer          'Formato YYYYMMDD 
    Public Property FechaFin As Integer             'Formato YYYYMMDD 
    Public Property FechaFirma As Integer           'Formato YYYYMMDD 
    Public Property Descripcion As String = ""      'Descripcion
    'Public Property Expediente As String = ""       'Expediente
    'Public Property Lote As String = ""             'Lote
    'Public Property ImporteLote As Double          'Importe lote
    Public Property Tipo As Integer                 '0->General, 1->Espífico
    Public Property Metodo As Integer               '0->Artículo, 1->Monetario
    Public Property Status As Integer               '0->Autorizado, 1-> En espera, 2->Borrador y 3->Terminado

    'Líneas
    Public Property Lineas As New List(Of EntAcuerdoLin)

    'SAP
    Public Property UserSAP As String = ""
    Public Property PassSAP As String = ""

End Class
