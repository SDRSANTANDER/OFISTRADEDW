Public Class EntProyectoSAP

    'DOCUWARE
    Public Property UserDW As String = ""           'Usuario Docuware
    Public Property PassDW As String = ""           'Password Docuware

    Public Property DOCIDDW As String = ""          'ID Docuware
    Public Property DOCURLDW As String = ""         'URL Docuware

    'INFORMACION
    Public Property NIFEmpresa As String = ""       'Identificador fiscal de la empresa
    Public Property Opcion As String = ""           'C->Crea documento, U->Actualiza datos docuware

    'PROYECTO
    Public Property Codigo As String = ""
    Public Property Nombre As String = ""
    Public Property FechaDesde As Integer
    Public Property FechaHasta As Integer
    Public Property Activo As String = ""           'S/N

    'SAP
    Public Property UserSAP As String = ""
    Public Property PassSAP As String = ""

End Class
