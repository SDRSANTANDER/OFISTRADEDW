Public Class EntActividadSAP

    'DOCUWARE
    Public Property UserDW As String = ""           'Usuario Docuware
    Public Property PassDW As String = ""           'Password Docuware

    Public Property DOCIDDW As String = ""          'ID Docuware
    Public Property DOCURLDW As String = ""         'URL Docuware

    'Generales
    Public Property NIFEmpresa As String = ""       'Identificador fiscal de la empresa
    Public Property Ambito As String = ""           'C->Compras, V->Ventas, I->Inventario
    Public Property Opcion As String = ""           'C->Crea documento, U->Actualiza datos docuware
    Public Property NIFTercero As String = ""       'Identificador fiscal de tercero 
    Public Property RazonSocial As String = ""      'Razón social proveedor (CardName)

    'Campos SAP
    Public Property Actividad As Integer            'Conversation=0, Meeting=1, Task=2, cn_Other=3, Note=4, Campaign=5
    Public Property Numero As String = ""
    Public Property Tipo As Integer
    Public Property Asunto As Integer
    Public Property AsignadoAUsuario As Integer
    Public Property AsignadoAEmpleado As Integer
    Public Property PersonaContacto As Integer
    'Public Property AsignadoPor As Integer
    Public Property Telefono As String = ""

    Public Property Comentarios As String = ""
    Public Property FechaHoraInicio As Long         'yyyyMMddHHmmss
    Public Property FechaHoraFin As Long            'yyyyMMddHHmmss

    Public Property Prioridad As Integer
    Public Property Emplazamiento As Integer

    Public Property Contenido As String = ""

    Public Property Cerrar As String = ""            'S/N

    'SAP
    Public Property UserSAP As String = ""
    Public Property PassSAP As String = ""

End Class
