Public Class EntCarta

    Public Property InvestigadorPrincipalId As Integer
    Public Property InvestigadorPrincipalNombre As String = ""
    Public Property InvestigadorPrincipalApellido As String = ""
    Public Property InvestigadorPrincipalMail As String = ""

    Public Property Numerador As String = ""
    Public Property DocumentoNumero As Integer

    Public Property SolicitudCompra As Integer
    Public Property PedidoVenta As Integer
    Public Property PedidoCompra As Integer

    Public Property InterlocutorId As String = ""
    Public Property InterLocutorRazonSocial As String = ""
    Public Property InterlocutorNumeroReferencia As String = ""

    Public Property FechaContable As Integer        'Formato yyyyMMdd
    Public Property FechaEntrega As Integer         'Formato yyyyMMdd

    Public Property InvestigadorId As Integer
    Public Property InvestigadorNombre As String = ""
    Public Property InvestigadorApellido As String = ""
    Public Property InvestigadorMail As String = ""

    Public Property LineaNumero As Integer
    Public Property LineaEstado As String = ""

    Public Property ArticuloId As String = ""
    Public Property ArticuloDescripcion As String = ""
    Public Property ArticuloReferencia As String = ""
    Public Property ArticuloGenericoLaboratorio As String = ""
    Public Property ArticuloTipoContrato As String = ""

    Public Property Cantidad As Double
    Public Property Precio As Double
    Public Property LineaTotal As Double

    Public Property ProyectoId As String = ""
    Public Property ProyectoDescripcion As String = ""
    Public Property Hito As String = ""
    Public Property NormaReparto As String = ""

    Public Property TextoLibre As String = ""
    Public Property CASS As String = ""

End Class



