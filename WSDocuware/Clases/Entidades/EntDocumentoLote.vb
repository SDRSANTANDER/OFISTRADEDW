Public Class EntDocumentoLote

    Public Property NumLote As String = ""
    Public Property Cantidad As Double

    Public Property Atributo1 As String = ""
    Public Property Atributo2 As String = ""

    Public Property FechaAdmision As Integer        'yyyyMMdd
    Public Property FechaFabricacion As Integer     'yyyyMMdd
    Public Property FechaVencimiento As Integer     'yyyyMMdd

    'Campos usuario
    Public Property CamposUsuario As New List(Of EntCampoUsuario)

End Class
