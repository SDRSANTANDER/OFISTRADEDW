Public Class EntLlamadaServicioSAP

    'DOCUWARE
    Public Property UserDW As String = ""            'Usuario Docuware
    Public Property PassDW As String = ""            'Password Docuware

    Public Property DOCIDDW As String = ""           'ID Docuware
    Public Property DOCURLDW As String = ""          'URL Docuware

    'Generales
    Public Property NIFEmpresa As String = ""        'Identificador fiscal de la empresa
    Public Property Ambito As String = ""            'C->Compras, V->Ventas, I->Inventario
    Public Property Opcion As String = ""            'C->Crea documento, U->Actualiza datos docuware
    Public Property NIFTercero As String = ""        'Identificador fiscal de tercero 
    Public Property RazonSocial As String = ""       'Razón social proveedor (CardName)

    'Campos SAP
    Public Property Serie As Integer
    Public Property Numero As String = ""
    Public Property Estado As Integer
    Public Property Prioridad As Integer

    Public Property PersonaContacto As Integer
    Public Property Telefono As String = ""
    Public Property NumAtCard As String = ""
    Public Property NumeroSerieFabricante As String = ""
    Public Property NumeroSerie As String = ""
    Public Property ItemCode As String = ""

    Public Property Asunto As String = ""

    Public Property Origen As Integer
    Public Property ProblemaTipo As Integer
    Public Property ProblemaSubtipo As Integer
    Public Property LlamadaTipo As Integer
    Public Property Tecnico As Integer

    Public Property TratadoPor As Integer

    Public Property Resolucion As String = ""


    'SAP
    Public Property UserSAP As String = ""
    Public Property PassSAP As String = ""

End Class
