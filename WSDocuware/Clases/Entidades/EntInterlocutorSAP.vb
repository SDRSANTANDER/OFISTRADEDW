Public Class EntInterlocutorSAP


    'DOCUWARE
    Public Property UserDW As String = ""            'Usuario Docuware
    Public Property PassDW As String = ""            'Password Docuware

    'Public Property DocBorradorDW As String = ""     'Documento borrador docuware (S/N)
    'Public Property CuentaContableDW As String = ""  'Cuenta contable docuware
    'Public Property ICReferenciaDW As String = ""   'IC referencia docuware (S/N)

    'Generales
    Public Property NIFEmpresa As String = ""
    Public Property Opcion As String = ""           'C->Crea artículo, U->Actualiza artículo

    'Campos SAP
    Public Property Serie As Integer
    Public Property Codigo As String = ""
    Public Property Nombre As String = ""
    Public Property NombreExtranjero As String = ""
    Public Property Tipo As String = ""             'Se considera si es numérico
    Public Property Grupo As String = ""            'Se considera si es numérico
    Public Property Moneda As String = ""
    Public Property NIFTercero As String = ""

    Public Property Telefono As String = ""
    Public Property Mail As String = ""
    Public Property Comentarios As String = ""

    Public Property Activo As String = ""           'S/N

    Public Property CamposUsuario As New List(Of EntCampoUsuario)

    'SAP
    Public Property UserSAP As String = ""
    Public Property PassSAP As String = ""


End Class
