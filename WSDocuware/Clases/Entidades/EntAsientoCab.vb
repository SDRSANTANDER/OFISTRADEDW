Public Class EntAsientoCab

    'DOCUWARE
    Public Property UserDW As String = ""               'Usuario Docuware
    Public Property PassDW As String = ""               'Password Docuware

    Public Property DOCIDDW As String = ""              'ID Docuware
    Public Property DOCURLDW As String = ""             'URL Docuware

    'CABECERA
    Public Property NIFEmpresa As String = ""           'Identificador fiscal de la empresa
    Public Property Opcion As String = ""               'C->Crea Asiento, U->Actualiza datos docuware
    Public Property Serie As Integer                    'Identificador serie
    'Public Property TransNum As String = ""             'Número de asiento emitido
    Public Property DocDate As Integer                  'Formato YYYYMMDD (Es la fecha de asiento, se mapea conta TaxDate)
    Public Property DocDateAux As Integer               'Formato YYYYMMDD (Es la fecha de asiento, se mapea conta TaxDate)
    Public Property AccountDate As Integer              'Formato YYYYMMDD (Es la fecha contable, se mapea contra RefDate)
    Public Property DocDueDate As Integer               'Formato YYYYMMDD

    Public Property Modelo As String = ""               'Código de módelo
    Public Property Indicador As String = ""            'Código de indicador
    Public Property Operacion As String = ""            'Código de operación
    Public Property Factura As String = ""              'Tipo de factura (F1, F2, R1, R2, R3, R4, R5, F3, F4, F5, F6, LC)

    Public Property Proyecto As String = ""             'Código de proyecto

    Public Property Ref1 As String = ""                 'Referencia 1
    Public Property Ref2 As String = ""                 'Referencia 2
    Public Property Ref3 As String = ""                 'Referencia 3

    Public Property Comments As String = ""             'Comentarios

    'BASE IMPONIBLE
    Public Property DocTotal As Double                  'Importe Asiento (impuestos incluidos)
    Public Property Currency As String = ""             'Moneda

    'DOCUMENTO DESTINO
    Public Property ObjTypeDestino As Integer           'Tipo de objeto destino según nomenclatura estándar SAP (30->Asiento...)

    'DOCUMENTO ORIGEN
    Public Property RefOrigen As String = ""            'Número de asiento
    Public Property TipoRefOrigen As String = ""        '"TransId", "TransNum"

    'SAP
    Public Property UserSAP As String = ""
    Public Property PassSAP As String = ""

    'Líneas
    Public Property Lineas As New List(Of EntAsientoLin)

    'Campos usuario
    Public Property CamposUsuario As New List(Of EntCampoUsuario)

End Class
