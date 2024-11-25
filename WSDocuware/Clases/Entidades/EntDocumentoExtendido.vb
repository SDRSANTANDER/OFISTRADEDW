Public Class EntDocumentoExtendido

    Public Property ObjectType As String = ""
    Public Property DocType As String = ""
    Public Property DocSerie As Integer
    Public Property DocEntry As Integer
    Public Property DocNum As Integer
    Public Property DocStatus As String = ""
    Public Property Cancelado As String = ""
    Public Property DocDate As Integer
    Public Property TaxDate As Integer
    Public Property DocDueDate As Integer
    Public Property TratadoDW As String = ""
    Public Property IdInterlocutor As String = ""
    Public Property RazonSocial As String = ""
    Public Property NIFTercero As String = ""
    Public Property DocRazonSocial As String = ""
    Public Property DocNIFTercero As String = ""
    Public Property NumAtCard As String = ""
    Public Property DocTotal As Double
    Public Property DocTotalFC As Double
    Public Property VatSum As Double
    Public Property VatSumFC As Double
    Public Property BaseTotal As Double
    Public Property BaseTotalFC As Double
    Public Property DiscPorc As Double
    Public Property DiscSum As Double
    Public Property DiscSumFC As Double
    Public Property DocRate As Double
    Public Property DocMoneda As String = ""
    Public Property CondicionPago As Integer
    Public Property ViaPago As String = ""
    Public Property EncargadoCV As Integer
    Public Property Titular As Integer
    Public Property IDDW As String = ""
    Public Property URLDW As String = ""
    Public Property Comentarios As String = ""

    'Líneas
    Public Property Lineas As New List(Of EntDocumentoExtendidoLin)

End Class
