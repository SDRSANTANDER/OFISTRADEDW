Public Class EntOportunidadCab

    'DOCUWARE
    Public Property UserDW As String = ""                   'Usuario Docuware
    Public Property PassDW As String = ""                   'Password Docuware

    Public Property DOCIDDW As String = ""                  'ID Docuware
    Public Property DOCURLDW As String = ""                 'URL Docuware

    'CABECERA
    Public Property NIFEmpresa As String = ""               'Identificador fiscal de la empresa
    Public Property Ambito As String = ""                   'C->Compras, V->Ventas
    Public Property Opcion As String = ""                   'C->Crea documento, U->Actualiza datos docuware
    Public Property NIFTercero As String = ""               'Identificador fiscal de tercero 
    Public Property RazonSocial As String = ""              'Razón social proveedor (CardName)
    Public Property Numero As String = ""                   'Número oportunidad
    Public Property Nombre As String = ""                   'Nombre oportunidad
    Public Property PersonaContacto As Integer              'Persona contacto
    Public Property EmpDptoCompraVenta As Integer           'Empleado departamento compras/ventas, si es 0 ninguno
    Public Property FechaInicio As Integer                  'Formato YYYYMMDD 

    'DOCUMENTO DESTINO
    Public Property ObjTypeDestino As Integer               'Tipo de objeto destino según nomenclatura estándar SAP (97->Oportunidades...)

    'CAMPOS POTENCIAL
    Public Property FechaCierrePrevista As Integer          'Fecha de cierre prevista
    Public Property ImportePotencial As Double              'Importe potencial en moneda local

    'RESUMEN
    Public Property Status As Integer                       '0->Abierta, 1->Ganada, 2->Perdida

    'SAP
    Public Property UserSAP As String = ""
    Public Property PassSAP As String = ""

End Class
