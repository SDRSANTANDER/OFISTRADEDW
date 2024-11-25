Public Class EntAsientoLin


    Public Property CuentaContable As String = ""       'Cuenta contable para documentos de tipo servicio

    Public Property ShortName As String = ""            'CardCode
    Public Property LineTipo As String = ""             'D-Debe, H-Haber
    Public Property LineTotal As Double             'Total de la línea, incluidos impuestos
    'Public Property BaseTotal As Double            'Base imponible de la línea, no usar

    'Public Property VATTotal As Double             'IVA de la línea
    Public Property VATGroup As String = ""         'Código del grupo de IVA
    'Public Property VATPorc As Double              'Porcentaje de IVA

    Public Property LineMemo As String = ""

    'Campos usuario
    Public Property CamposUsuario As New List(Of EntCampoUsuario)

End Class
