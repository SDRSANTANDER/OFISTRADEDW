Public Class EntDocumentoCab

    'DOCUWARE
    Public Property UserDW As String = ""               'Usuario Docuware
    Public Property PassDW As String = ""               'Password Docuware

    Public Property DOCIDDW As String = ""              'ID Docuware
    Public Property DOCURLDW As String = ""             'URL Docuware

    'CABECERA
    Public Property NIFEmpresa As String = ""           'Identificador fiscal de la empresa
    Public Property Ambito As String = ""               'C->Compras, V->Ventas, I->Inventario
    Public Property Opcion As String = ""               'C->Crea documento, U->Actualiza datos docuware
    Public Property Serie As Integer                    'Identificador serie
    Public Property NIFTercero As String = ""           'Identificador fiscal de tercero 
    Public Property RazonSocial As String = ""          'Razón social proveedor (CardName)
    'Public Property DocNum As String = ""               'Número de documento emitido
    Public Property DocDate As Integer                  'Formato YYYYMMDD (Es la fecha de documento, se mapea conta TaxDate)
    Public Property DocDateAux As Integer               'Formato YYYYMMDD (Es la fecha de documento, se mapea conta TaxDate)
    Public Property AccountDate As Integer              'Formato YYYYMMDD (Es la fecha contable, se mapea contra DocDate)
    Public Property DocDueDate As Integer               'Formato YYYYMMDD

    Public Property NumAtCard As String = ""            'Referencia cliente
    Public Property Comments As String = ""             'Comentarios

    Public Property JournalMemo As String = ""          'Entrada diario

    Public Property BloqueoPago As Integer              'Tipo de bloqueo de pago, si es 0 ninguno

    Public Property Sucursal As Integer                 'Sucursal, si es 0 ninguna

    Public Property Proyecto As String = ""             'Código de proyecto

    Public Property ClaseExpedicion As Integer          'Clase expedición, si es 0 ninguna
    Public Property Responsable As String = ""          'Responsable
    Public Property Titular As Integer                  'Titular, si es 0 ninguno
    Public Property EmpDptoCompraVenta As Integer       'Empleado departamento compras/ventas, si es 0 ninguno

    'Public Property PortesTipo As Integer              'Portes código del documento
    'Public Property Portes As Double                   'Portes importe del documento

    'Public Property CopiarNormal As String = ""        'Forma de copia S -> normal, N -> especificada por el cliente (Catalysis)
    Public Property ControlDiferencia As Double         'Valor máximo de diferéncia entre DocTotal de SAP y DocTotal DW


    'DOCUMENTO DESTINO
    Public Property ObjTypeDestino As Integer           'Tipo de objeto destino según nomenclatura estándar SAP (13->Facturas, 15->Entregas, 22->Pedidos...)
    Public Property RefDestino As String                'Número de afactura destino
    Public Property TipoRefDestino As String            '"DocNum","DocEntry" o "NumAtCard"
    Public Property DocType As String = ""              'Tipo de documento (S->Servicios, I->Artículos)
    Public Property Draft As String = ""                'S->Borrador, N->Documento definitivo, I->Configuración interlocutor
    Public Property Reserva As String = ""              'Factura de reserva S/N


    'DOCUMENTO ORIGEN
    Public Property ObjTypeOrigen As Integer            'Tipo de objeto origen según nomenclatura estándar SAP (0->Por defecto, 13->Facturas, 15->Entregas, 22->Pedidos...)
    Public Property RefOrigen As String = ""            'Número de albaranes/pedidos origen
    Public Property TipoRefOrigen As String = ""        '"DocNum","DocEntry" o "NumAtCard"


    'FACTURAS ANTICIPO
    Public Property AnticipoRefOrigen As String = ""             'Número de facturas anticipo origen
    Public Property AnticipoTipoRefOrigen As String = ""         '"DocNum","DocEntry" o "NumAtCard"


    'BASE IMPONIBLE
    Public Property DocTotal As Double              'Importe documento (impuestos incluidos)
    Public Property IRPFImporte As Double           'Importe IRPF
    'Public Property IRPFCuota As Double            'Cuota IRPF
    Public Property Currency As String = ""         'Moneda


    'CAMPOS ANTICIPO
    Public Property PorcentajeAnticipo As Double    'Porcentaje de anticipo, no usar (pasar 0)


    'CAMPOS SOLICITUDES DE COMPRA
    Public Property Solicitante As String = ""      'Empleado que solicita la compra
    Public Property Email As String = ""            'Mail del empleado que solicita la compra


    'CAMPOS DIRECCIONES
    Public Property DireccionEnvioCodigo As String = ""         'Direccion envío
    Public Property DireccionEnvioDetalle As String = ""        'Direccion envío

    Public Property DireccionFacturaCodigo As String = ""       'Direccion factura
    Public Property DireccionFacturaDetalle As String = ""      'Direccion factura


    'CAMPOS CONTACTO
    Public Property ContactoCodigo As Integer           'Contacto código
    Public Property ContactoEmail As String = ""        'Email contacto


    'ENTREGA/PRECIO
    Public Property AduanaImporte As Double             'Importe aduana


    'CAMPOS COBRO/PAGO
    Public Property CobroPagoTipo As String = ""        'Efectivo/Efecto/Transferencia/Cheque/Tarjeta
    Public Property CobroPagoImporte As Double
    Public Property CobroPagoCuenta As String = ""
    Public Property CobroPagoReferencia As String = ""
    Public Property CobroPagoViaPago As String = ""
    Public Property CobroPagoBanco As String = ""
    Public Property CobroPagoPais As String = ""
    Public Property CobroPagoNumCheque As Integer
    Public Property CobroPagoTarjeta As Integer
    'Public Property CobroPagoNumTarjeta As String = ""
    Public Property CobroPagoValidezTarjeta As Integer
    Public Property CobroPagoComprobanteTarjeta As String = ""
    Public Property CobroPagoProyecto As String = ""

    'Finanzas
    Public Property FinanzasNIF As String = ""
    Public Property FinanzasRazon As String = ""

    'SAP
    Public Property UserSAP As String = ""
    Public Property PassSAP As String = ""

    'Líneas
    Public Property Lineas As New List(Of EntDocumentoLin)

    'Portes
    Public Property PortesDetalle As New List(Of EntDocumentoPorte)

    'Documentos relacionados
    Public Property DocRelacionados As New List(Of EntDocumentoRelacionado)

    'Campos usuario
    Public Property CamposUsuario As New List(Of EntCampoUsuario)

End Class
