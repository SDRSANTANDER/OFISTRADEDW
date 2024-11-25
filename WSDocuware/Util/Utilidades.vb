Imports SAPbobsCOM
Imports System.Reflection

Public Class Utilidades

#Region "Sociedades"

    'Sociedades
    Public Enum eSociedad
        SOCIEDAD1 = 1
    End Enum

    Public Shared ReadOnly Property NOMBRESOCIEDAD(ByVal Sociedad As eSociedad) As String

        Get

            If Sociedad = eSociedad.SOCIEDAD1 Then
                Return "SOCIEDAD1"
            Else
                Return ""
            End If

        End Get

    End Property

    Public Shared ReadOnly Property SOCIEDADPORNIF(ByVal NIF As String) As eSociedad

        Get

            If NIF = "ESB59806794" Then
                Return eSociedad.SOCIEDAD1
            Else
                Throw New Exception("Sociedad por NIF no definida (" & NIF & ")")
            End If

        End Get

    End Property

#End Region

#Region "Variables"

    Public Const DBTypeHANA As String = "9"

    Public Const ArticuloLineaEspecialTexto As String = "LINEAESPECIALTEXTOSAP"
    Public Const ArticuloLineaEspecialSubtotal As String = "LINEAESPECIALSUBTOTALSAP"

    'Código de respuesta
    Public Structure Respuesta
        Const Ok = "Ok"
        Const Ko = "Ko"
    End Structure

    'Si/No
    Public Structure SN
        Const Si = "S"
        Const No = "N"
        Const Yes = "Y"
    End Structure

    'Búsqueda por IC
    Public Structure BusquedaIC
        Const CardCode = "CardCode"
        Const NIF = "NIF"
        Const U_SEIICDW = "U_SEIICDW"
    End Structure

    'Definición de categorías de IVA
    Public Structure IVA
        Const Soportado = "I"
        Const Repercutido = "O"
    End Structure

    'Ambito
    Public Structure Ambito
        Const Compras = "C"
        Const Ventas = "V"
        Const Inventario = "I"
    End Structure

    'Definición de cómo crear el documento
    Public Structure Draft
        Const Borrador = "S"
        Const Firme = "N"
        Const Interlocutor = "I"
    End Structure

    'Definición de tipo de documento
    Public Structure DocType
        Const Servicio = "S"
        Const Articulo = "I"
        Const Cuenta = "A"
    End Structure

    'Definición de estado de documento
    Public Structure DocStatus
        Const Abierto = "O"
        Const Cerrado = "C"
    End Structure

    'Definición de tipo de IC
    Public Structure CardType
        Const Cliente = "C"
        Const Proveedor = "S"
        Const Lead = "L"
        Const ClienteConPedidosAbiertos = "CPA"
    End Structure


    'Definición de tipo de artículo
    Public Structure ItemType
        Const Articulo = "0"
        Const Trabajo = "1"
        Const Viaje = "2"
        Const ActivoFijo = "3"
    End Structure

    'Opción (crear o actualizar)
    Public Structure Opcion
        Const Crear = "C"
        Const Actualizar = "U"
    End Structure

    'Tipos generación
    Public Enum Generar
        CONTRATOMENOR = 1
    End Enum

    'StatusImpuesto
    Public Structure StatusImpuesto
        Const Adquisiciones = "E"
        Const Obligatorio = "Y"
        Const Exento = "N"
    End Structure

    'Pagos/Cobros 
    Public Structure CobroPago
        Const Efectivo As String = "Efectivo"
        Const Efecto As String = "Efecto"
        Const Transferencia As String = "Transferencia"
        Const Cheque As String = "Cheque"
        Const Tarjeta As String = "Tarjeta"
    End Structure

    'Asientos
    Public Structure AsientoLineaTipo
        Const Debe = "D"
        Const Haber = "H"
    End Structure

    'Obtencion de datos
    Public Enum Peticion
        Empleados = 1
        Interlocutores = 2
        Articulos = 3
        CentrosCoste = 4
        Proyectos = 5
        DocumentosAbiertos = 6
        ICIQ_SolicitudesPedidosCompraIP = 7
        ICIQ_RelacionesDocumentosCompra = 8
        Sucursales = 9
        OrigenDocumentos = 10
        Series = 11
        GruposArticulo = 12
        GruposUnidadesMedida = 13
        Fabricantes = 14
        ClasesExpedicion = 15
        GruposImpositivos = 16
        ValoresValidos = 17
        TarjetasBanco = 18
        ViasPago = 19
        Monedas = 20
        TasasCambio = 21
        'Nuevas
        ActividadesTipo = 22
        ActividadesAsunto = 23
        ActividadesEmplazamiento = 24
        LlamadasServicioEstado = 25
        LlamadasServicioTipo = 26
        LlamadasServicioOrigen = 27
        LlamadasServicioProblemaTipo = 28
        LlamadasServicioProblemaSubtipo = 29
        Almacenes = 30
        GruposInterlocutor = 31
        AURIA_CentrosCosteRIC = 32
        CondicionesPago = 33
        Responsables = 34
        EmpleadosDptoCV = 35
        BancosPropios = 36
        InterlocutoresBancos = 37
        Portes = 38
        GrupoLaPuente_PedidosCompraLineas = 39
        InterlocutoresDirecciones = 40
        InterlocutoresContactos = 41
        FacturasAnticipo = 42
        ValoresTablas = 43
        DocumentosExtendida = 44
        TarjetasEquipos = 45
        AsientosModelos = 46
        AsientosIndicadores = 47
        AsientosOperaciones = 48
        Idiomas = 49
        CuentasContables = 50
        AcuerdosGlobales = 51
        NormasReparto = 52
        'Activos fijos
        ActivosFijosClases = 53
        ActivosFijosGrupos = 54
        ActivosFijosGruposAmortizacion = 55
        ActivosFijosEmplazamientos = 56
        ActivosFijosAreasValoracion = 57
        LlamadasServicio = 58
        Actividades = 59
        LlamadasServicioCola = 60
        PreciosEntregaAbiertos = 61
        Ceamsa_FacturacionConceptoEstandar = 62
        Ceamsa_EmpleadosExtendido = 63
        VHIO_Documentos = 64
        ICIQ_DocumentosProyectos = 65
    End Enum

    'Obtencion de campos
    Public Enum Campos
        NumDocumento = 1
        Importe = 2
        FechaVencimiento = 3
        ResponsableMail = 4
        Comentarios = 5
        FechaImporteVencimiento = 6
        PagadoVencimiento = 7
        ImporteSinIVA = 8
        ViaPago = 9
        Moneda = 10
        NumTransaccion = 11
        Proyecto = 12
        CampoUsuario = 13
        Sucursal = 14
        CentroCoste = 15
        NumEnvio = 16
        NumDestino = 17
        Titular = 18
    End Enum

    'Referencia origen 
    Public Structure RefOrigen
        Const DocEntry = "DocEntry"
        Const DocNum = "DocNum"
        Const NumAtCard = "NumAtCard"
        Const TransId = "TransId"
        Const TransNum = "TransNum"
        Const TransRef1 = "TransRef1"
        Const TransRef2 = "TransRef2"
        Const TransRef3 = "TransRef3"
    End Structure

    'ObjType 
    Public Structure ObjType
        Const PedidoCompra = 22
        Const EntregaCompra = 20
        Const FacturaCompra = 18
        Const FacturaCompraAnticipo = 204
        Const AbonoCompra = 19
        Const PedidoVenta = 17
        Const ReciboProduccion = 59
        Const OrdenProduccion = 202
        Const SolicitudCompra = 1470000113
        Const SolicitudPedidoCompra = 540000006
        Const Cobro = 24
        Const Pago = 46
        Const Traslado = 67
        Const EntradaMercancias = 59
        Const SalidaMercancias = 60
        Const PrecioEntrega = 69
        Const LlamadaServicio = 191
        Const Actividad = 33
        Const Oportunidad = 97
        Const AcuerdoGlobal = 1250000025
        Const Asiento = 30
    End Structure

    'Tablas 
    Public Structure TablaSAP
        Const Lotes As String = "OBTN"
    End Structure

    'CampoTipo 
    Public Structure CampoTipo
        Const Int As String = "INT"
        Const Dec As String = "DEC"
        Const Txt As String = "TXT"
    End Structure

#End Region

#Region "Funciones"

    Public Shared Function getTablaDeObjType(ByVal ObjectType As Integer) As String

        'Devuelve la tabla que hay que usar 
        Dim retVal As String = ""

        Try

            Select Case ObjectType

                Case 1
                    retVal = "OACT"
                Case 2
                    retVal = "OCRD"
                Case 3
                    retVal = "ODSC"
                Case 4
                    retVal = "OITM"
                Case 5
                    retVal = "OVTG"
                Case 6
                    retVal = "OPLN"
                Case 7
                    retVal = "OSPP"
                Case 8
                    retVal = "OITG"
                Case 9
                    retVal = "ORTM"
                Case 10
                    retVal = "OCRG"
                Case 11
                    retVal = "OCPR"
                Case 12
                    retVal = "OUSR"
                Case 13
                    retVal = "OINV"
                Case 14
                    retVal = "ORIN"
                Case 15, 1517
                    retVal = "ODLN"
                Case 16
                    retVal = "ORDN"
                Case 17
                    retVal = "ORDR"
                Case 18
                    retVal = "OPCH"
                Case 19
                    retVal = "ORPC"
                Case 20, 2022
                    retVal = "OPDN"
                Case 21
                    retVal = "ORPD"
                Case 22
                    retVal = "OPOR"
                Case 23
                    retVal = "OQUT"
                Case 24
                    retVal = "ORCT"
                Case 25
                    retVal = "ODPS"
                Case 26
                    retVal = "OMTH"
                Case 27
                    retVal = "OCHH"
                Case 28
                    retVal = "OBTF"
                Case 29
                    retVal = "OBTD"
                Case 30
                    retVal = "OJDT"
                Case 31
                    retVal = "OITW"
                Case 32
                    retVal = "OADP"
                Case 33
                    retVal = "OCLG"
                Case 34
                    retVal = "ORCR"
                Case 35
                    retVal = "ONNM"
                Case 36
                    retVal = "OCRC"
                Case 37
                    retVal = "OCRN"
                Case 38
                    retVal = "OIDX"
                Case 39
                    retVal = "OADM"
                Case 40
                    retVal = "OCTG"
                Case 41
                    retVal = "OPRF"
                Case 42
                    retVal = "OBNK"
                Case 43
                    retVal = "OMRC"
                Case 44
                    retVal = "OCQG"
                Case 45
                    retVal = "OTRC"
                Case 46
                    retVal = "OVPM"
                Case 47
                    retVal = "OSRL"
                Case 48
                    retVal = "OALC"
                Case 49
                    retVal = "OSHP"
                Case 50
                    retVal = "OLGT"
                Case 51
                    retVal = "OWGT"
                Case 52
                    retVal = "OITB"
                Case 53
                    retVal = "OSLP"
                Case 54
                    retVal = "OFLT"
                Case 55
                    retVal = "OTRT"
                Case 56
                    retVal = "OARG"
                Case 57
                    retVal = "OCHO"
                Case 58
                    retVal = "OINM"
                Case 59
                    retVal = "OIGN"
                Case 60
                    retVal = "OIGE"
                Case 61
                    retVal = "OPRC"
                Case 62
                    retVal = "OOCR"
                Case 63
                    retVal = "OPRJ"
                Case 64
                    retVal = "OWHS"
                Case 65
                    retVal = "OCOG"
                Case 66
                    retVal = "OITT"
                Case 67
                    retVal = "OWTR"
                Case 68
                    retVal = "OWKO"
                Case 69
                    retVal = "OIPF"
                Case 70
                    retVal = "OCRP"
                Case 71
                    retVal = "OCDT"
                Case 72
                    retVal = "OCRH"
                Case 73
                    retVal = "OSCN"
                Case 74
                    retVal = "OCRV"
                Case 75
                    retVal = "ORTT"
                Case 76
                    retVal = "ODPT"
                Case 77
                    retVal = "OBGT"
                Case 78
                    retVal = "OBGD"
                Case 79
                    retVal = "ORCN"
                Case 80
                    retVal = "OALT"
                Case 81
                    retVal = "OALR"
                Case 82
                    retVal = "OAIB"
                Case 83
                    retVal = "OAOB"
                Case 84
                    retVal = "OCLS"
                Case 85
                    retVal = "OSPG"
                Case 86
                    retVal = "SPRG"
                Case 87
                    retVal = "OMLS"
                Case 88
                    retVal = "OENT"
                Case 89
                    retVal = "OSAL"
                Case 90
                    retVal = "OTRA"
                Case 91
                    retVal = "OBGS"
                Case 92
                    retVal = "OIRT"
                Case 93
                    retVal = "OUDG"
                Case 94
                    retVal = "OSRI"
                Case 95
                    retVal = "OFRT"
                Case 96
                    retVal = "OFRC"
                Case 97
                    retVal = "OOPR"
                Case 98
                    retVal = "OOIN"
                Case 99
                    retVal = "OOIR"
                Case 100
                    retVal = "OOSR"
                Case 101
                    retVal = "OOST"
                Case 102
                    retVal = "OOFR"
                Case 103
                    retVal = "OCLT"
                Case 104
                    retVal = "OCLO"
                Case 105
                    retVal = "OISR"
                Case 106
                    retVal = "OIBT"
                Case 107
                    retVal = "OALI"
                Case 108
                    retVal = "OPRT"
                Case 109
                    retVal = "OCMT"
                Case 110
                    retVal = "OUVV"
                Case 111
                    retVal = "OFPR"
                Case 112
                    retVal = "ODRF"
                Case 113
                    retVal = "OSRD"
                Case 114
                    retVal = "OUDC"
                Case 115
                    retVal = "OPVL"
                Case 116
                    retVal = "ODDT"
                Case 117
                    retVal = "ODDG"
                Case 118
                    retVal = "OUBR"
                Case 119
                    retVal = "OUDP"
                Case 120
                    retVal = "OWST"
                Case 121
                    retVal = "OWTM"
                Case 122
                    retVal = "OWDD"
                Case 123
                    retVal = "OCHD"
                Case 124
                    retVal = "CINF"
                Case 125
                    retVal = "OEXD"
                Case 126
                    retVal = "OSTA"
                Case 127
                    retVal = "OSTT"
                Case 128
                    retVal = "OSTC"
                Case 129
                    retVal = "OCRY"
                Case 130
                    retVal = "OCST"
                Case 131
                    retVal = "OADF"
                Case 132
                    retVal = "OCIN"
                Case 133
                    retVal = "OCDC"
                Case 134
                    retVal = "OQCN"
                Case 135
                    retVal = "OIND"
                Case 136
                    retVal = "ODMW"
                Case 137
                    retVal = "OCSTN"
                Case 138
                    retVal = "OIDC"
                Case 139
                    retVal = "OGSP"
                Case 140
                    retVal = "OPDF"
                Case 141
                    retVal = "OQWZ"
                Case 142
                    retVal = "OASG"
                Case 143
                    retVal = "OASC"
                Case 144
                    retVal = "OLCT"
                Case 145
                    retVal = "OTNN"
                Case 146
                    retVal = "OCYC"
                Case 147
                    retVal = "OPYM"
                Case 148
                    retVal = "OTOB"
                Case 149
                    retVal = "ORIT"
                Case 150
                    retVal = "OBPP"
                Case 151
                    retVal = "ODUN"
                Case 152
                    retVal = "CUFD"
                Case 153
                    retVal = "OUTB"
                Case 154
                    retVal = "OCUMI"
                Case 155
                    retVal = "OPYD"
                Case 156
                    retVal = "OPKL"
                Case 157
                    retVal = "OPWZ"
                Case 158
                    retVal = "OPEX"
                Case 159
                    retVal = "OPYB"
                Case 160
                    retVal = "OUQR"
                Case 161
                    retVal = "OCBI"
                Case 162
                    retVal = "OMRV"
                Case 163
                    retVal = "OCPI"
                Case 164
                    retVal = "OCPV"
                Case 165
                    retVal = "OCSI"
                Case 166
                    retVal = "OCSV"
                Case 167
                    retVal = "OSCS"
                Case 168
                    retVal = "OSCT"
                Case 169
                    retVal = "OSCP"
                Case 170
                    retVal = "OCTT"
                Case 171
                    retVal = "OHEM"
                Case 172
                    retVal = "OHTY"
                Case 173
                    retVal = "OHST"
                Case 174
                    retVal = "OHTR"
                Case 175
                    retVal = "OHED"
                Case 176
                    retVal = "OINS"
                Case 177
                    retVal = "OAGP"
                Case 178
                    retVal = "OWHT"
                Case 179
                    retVal = "ORFL"
                Case 180
                    retVal = "OVTR"
                Case 181
                    retVal = "OBOE"
                Case 182
                    retVal = "OBOT"
                Case 183
                    retVal = "OFRM"
                Case 184
                    retVal = "OPID"
                Case 185
                    retVal = "ODOR"
                Case 186
                    retVal = "OHLD"
                Case 187
                    retVal = "OCRB"
                Case 188
                    retVal = "OSST"
                Case 189
                    retVal = "OSLT"
                Case 190
                    retVal = "OCTR"
                Case 191
                    retVal = "OSCL"
                Case 192
                    retVal = "OSCO"
                Case 193
                    retVal = "OUKD"
                Case 194
                    retVal = "OQUE"
                Case 195
                    retVal = "OIWZ"
                Case 196
                    retVal = "ODUT"
                Case 197
                    retVal = "ODWZ"
                Case 198
                    retVal = "OFCT"
                Case 199
                    retVal = "OMSN"
                Case 200
                    retVal = "OTER"
                Case 201
                    retVal = "OOND"
                Case 202
                    retVal = "OWOR"
                Case 203
                    retVal = "ODPI"
                Case 204
                    retVal = "ODPO"
                Case 205
                    retVal = "OPKG"
                Case 206
                    retVal = "OUDO"
                Case 207
                    retVal = "ODOW"
                Case 208
                    retVal = "ODOX"
                Case 209
                    retVal = ""
                Case 210
                    retVal = "OHPS"
                Case 211
                    retVal = "OHTM"
                Case 212
                    retVal = "OORL"
                Case 213
                    retVal = "ORCM"
                Case 214
                    retVal = "OUPT"
                Case 215
                    retVal = "OPDT"
                Case 216
                    retVal = "OBOX"
                Case 217
                    retVal = "OCLA"
                Case 218
                    retVal = "OCHF"
                Case 219
                    retVal = "OCSHS"
                Case 220
                    retVal = "OACP"
                Case 221
                    retVal = "OATC"
                Case 222
                    retVal = "OGFL"
                Case 223
                    retVal = "OLNG"
                Case 224
                    retVal = "OMLT"
                Case 225
                    retVal = "OAPA3"
                Case 226
                    retVal = "OAPA4"
                Case 227
                    retVal = "OAPA5"
                Case 229
                    retVal = "SDIS"
                Case 230
                    retVal = "OSVR"
                Case 231
                    retVal = "DSC1"
                Case 232
                    retVal = "RDOC"
                Case 233
                    retVal = "ODGP"
                Case 234
                    retVal = "OMHD"
                Case 238
                    retVal = "OACG"
                Case 239
                    retVal = "OBCA"
                Case 241
                    retVal = "OCFT"
                Case 242
                    retVal = "OCFW"
                Case 247
                    retVal = "OBPL"
                Case 250
                    retVal = "OJPE"
                Case 251
                    retVal = "ODIM"
                Case 254
                    retVal = "OSCD"
                Case 255
                    retVal = "OSGP"
                Case 256
                    retVal = "OMGP"
                Case 257
                    retVal = "ONCM"
                Case 258
                    retVal = "OCFP"
                Case 259
                    retVal = "OTSC"
                Case 260
                    retVal = "OUSG"
                Case 261
                    retVal = "OCDP"
                Case 263
                    retVal = "ONFN"
                Case 264
                    retVal = "ONFT"
                Case 265
                    retVal = "OCNT"
                Case 266
                    retVal = "OTCD"
                Case 267
                    retVal = "ODTY"
                Case 268
                    retVal = "OPTF"
                Case 269
                    retVal = "OIST"
                Case 271
                    retVal = "OTPS"
                Case 275
                    retVal = "OTFC"
                Case 276
                    retVal = "OFML"
                Case 278
                    retVal = "OCNA"
                Case 280
                    retVal = "OTSI"
                Case 281
                    retVal = "OTPI"
                Case 283
                    retVal = "OCCD"
                Case 290
                    retVal = "ORSC"
                Case 291
                    retVal = "ORSG"
                Case 292
                    retVal = "ORSB"
                Case 300
                    retVal = ""
                Case 305
                    retVal = ""
                Case 321
                    retVal = "OITR"
                Case 541
                    retVal = "OPOS"
                Case 1179
                    retVal = "ODRF"
                Case 10000105
                    retVal = "OMSG"
                Case 10000044
                    retVal = "OBTN"
                Case 10000045
                    retVal = "OSRN"
                Case 10000062
                    retVal = "OIVK"
                Case 10000071
                    retVal = "OIQR"
                Case 10000073
                    retVal = "OFYM"
                Case 10000074
                    retVal = "OSEC"
                Case 10000075
                    retVal = "OCSN"
                Case 10000077
                    retVal = "ONOA"
                Case 10000196
                    retVal = "RTYP"
                Case 10000197
                    retVal = "OUGP"
                Case 10000199
                    retVal = "OUOM"
                Case 10000203
                    retVal = "OBFC"
                Case 10000204
                    retVal = "OBAT"
                Case 10000205
                    retVal = "OBSL"
                Case 10000206
                    retVal = "OBIN"
                Case 140000041
                    retVal = "ODNF"
                Case 231000000
                    retVal = "OUGR"
                Case 234000004
                    retVal = "OEGP"
                Case 243000001
                    retVal = "OGPC"
                Case 310000001
                    retVal = "OIQI"
                Case 310000008
                    retVal = "OBTW"
                Case 410000005
                    retVal = "OLLF"
                Case 480000001
                    retVal = "OHET"
                Case 540000005
                    retVal = "OTCX"
                Case 540000006
                    retVal = "OPQT"
                Case 540000040
                    retVal = "ORCP"
                Case 540000042
                    retVal = "OCCT"
                Case 540000048
                    retVal = "OACR"
                Case 540000056
                    retVal = "ONFM"
                Case 540000067
                    retVal = "OBFI"
                Case 540000068
                    retVal = "OBBI"
                Case 1210000000
                    retVal = "OCPT"
                Case 1250000001
                    retVal = "OWTQ"
                Case 1250000025
                    retVal = "OOAT"
                Case 1320000000
                    retVal = "OKPI"
                Case 1320000002
                    retVal = "OTGG"
                Case 1320000012
                    retVal = "OCPN"
                Case 1320000028
                    retVal = "OROC"
                Case 1320000039
                    retVal = "OPSC"
                Case 1470000000
                    retVal = "ODTP"
                Case 1470000002
                    retVal = "OADT"
                Case 1470000003
                    retVal = "ODPA"
                Case 1470000004
                    retVal = "ODPP"
                Case 1470000032
                    retVal = "OACS"
                Case 1470000046
                    retVal = "OAGS"
                Case 1470000048
                    retVal = "ODMC"
                Case 1470000049
                    retVal = "OACQ"
                Case 1470000057
                    retVal = "OGAR"
                Case 1470000060
                    retVal = "OACD"
                Case 1470000062
                    retVal = "OBCD"
                Case 1470000065
                    retVal = "OINC"
                Case 1470000077
                    retVal = "OEDG"
                Case 1470000092
                    retVal = "OCCS"
                Case 1470000113
                    retVal = "OPRQ"
                Case 1620000000
                    retVal = "OWLS"
                Case Else
                    retVal = ""

            End Select

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Shared Function getOrigenDeObjTypeNoDirecto(ByVal ObjectType As Integer) As KeyValuePair(Of Integer, Integer)

        'Devuelve los ObjType que hay que usar en orígenes no directos
        Dim retVal As New KeyValuePair(Of Integer, Integer)

        Try

            Select Case ObjectType

                Case 2022
                    retVal = New KeyValuePair(Of Integer, Integer)(20, 22)

                Case 1517
                    retVal = New KeyValuePair(Of Integer, Integer)(15, 17)

                Case Else
                    retVal = New KeyValuePair(Of Integer, Integer)

            End Select

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Shared Function getEsDocumentoSolicitud(ByVal ObjectType As Integer) As Boolean

        'Devuelve si el documento es de tipo solicitud
        Dim retVal As Boolean = False

        Try

            Select Case ObjectType

                'Solicitud pedido de compra
                Case "1470000113"
                    retVal = True
                    'Case "540000006"
                    '    retVal = True

            End Select

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Shared Function getEsDocumentoNoDirecto(ByVal ObjectType As Integer) As Boolean

        'Devuelve si el documento es de tipo directo o no
        Dim retVal As Boolean = False

        Try

            Select Case ObjectType

                'Albaranes de compra/venta desde pedidos
                Case 2022, 1517
                    retVal = True

            End Select

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Shared Function getEsDocumentoVenta(ByVal ObjectType As Integer) As Boolean

        'Devuelve si el documento es de tipo venta
        Dim retVal As Boolean = False

        Try

            Select Case ObjectType

                Case 13, 14, 15, 16, 17, 23, 24, 203, 234000031
                    retVal = True

            End Select

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Shared Function getEsFactura(ByVal ObjectType As Integer) As Boolean

        'Devuelve si el documento es factura de venta o compra
        Dim retVal As Boolean = False

        Try

            Select Case ObjectType

                Case 13, 18
                    retVal = True

            End Select

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Shared Function getComprobarVencimientos(ByVal ObjectType As Integer) As Boolean

        'Devuelve si el documento tiene vencimientos
        Dim retVal As Boolean = False

        Try

            Select Case ObjectType

                Case 13, 18
                    retVal = True

            End Select

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Shared Function getTablaAnticipoDeObjType(ByVal ObjectType As Integer) As String

        'Devuelve la tabla que hay que usar 
        Dim retVal As String = ""

        Try

            Select Case ObjectType

                Case 13
                    retVal = "ODPI"

                Case 18
                    retVal = "ODPO"

                Case Else
                    retVal = ""

            End Select

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Shared Function getEsCobroPago(ByVal Tabla As String) As Boolean

        'Devuelve si el documento es de tipo cobro/pago
        Dim retVal As Boolean = False

        Try

            Select Case Tabla

                Case "ORCT", "OVPM"
                    retVal = True

            End Select

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Shared Function getBaseObjTypePrecioEntrega(ByVal ObjectType As Integer) As String

        'Devuelve el object type que hay que usar
        Dim retVal As LandedCostBaseDocumentTypeEnum = LandedCostBaseDocumentTypeEnum.asEmpty

        Try

            Select Case ObjectType

                Case 18
                    retVal = LandedCostBaseDocumentTypeEnum.asPurchaseInvoice

                Case 20
                    retVal = LandedCostBaseDocumentTypeEnum.asGoodsReceiptPO

                Case 69
                    retVal = LandedCostBaseDocumentTypeEnum.asLandedCosts

                Case Else

                    retVal = LandedCostBaseDocumentTypeEnum.asDefault

            End Select

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Shared Function ComprobarCredenciales(ByVal User As String, ByVal Pass As String) As Boolean

        Dim retVal As Boolean = False

        Try

            If User = ConfigurationManager.AppSettings.Item("USERDW").ToString AndAlso Pass = ConfigurationManager.AppSettings.Item("PASSDW").ToString Then
                retVal = True
            Else
                Throw New Exception("Usuario o Password incorrecto (" & User & " - " & Pass & ")")
            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw ex
        End Try

        Return retVal

    End Function

    Public Shared Function ObtenerGrupoIVA(ByVal sGrupoIVA As String, ByVal sNumero As String) As String

        'Devuelve el grupo de IVA cambiando el número
        Dim retVal As String = ""

        Try

            'retVal = Regex.Replace(sGrupoIVA, "[0-9]", sNumero)
            retVal = Regex.Replace(sGrupoIVA, "[0-9]", "") & sNumero

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw ex
        End Try

        Return retVal

    End Function

    Public Shared Function MensajeSalida(ByVal oResultado As EntResultado) As String

        'Devolver resultado como CODIGO:X#MENSAJE:Y# 
        Dim retVal As String = ""

        Try

            retVal &= "CODIGO:" & oResultado.CODIGO & "#"
            retVal &= "MENSAJE:" & oResultado.MENSAJE & "#"
            retVal &= "MENSAJEAUX:" & oResultado.MENSAJEAUX & "#"

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw ex
        End Try

        Return retVal

    End Function

    Public Shared Function MensajeSalida(ByVal strJson As String) As String

        'Devolver resultado como CODIGO:X#MENSAJE:Y|
        Dim retVal As String = ""

        Try

            retVal = strJson.Replace("{", "").Replace("}", "").Replace(""",""", """|""").Replace("""", "") & "|"

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw ex
        End Try

        Return retVal

    End Function

    Public Shared Function MensajeAuxSalida(ByVal strJson As String) As String

        'Devolver resultado sin contrbarras
        Dim retVal As String = ""

        Try

            retVal = strJson.Replace("\""", """").Replace("""\", """")

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw ex
        End Try

        Return retVal

    End Function

#End Region

#Region "Formato campos"

    Public Shared Function getDoubleFromString(ByVal str As String) As Double

        Dim retval As Double = 0

        If Not String.IsNullOrEmpty(str) Then
            retval = CDbl(str.Replace(".", ","))
        End If

        Return retval

    End Function

#End Region

#Region "Formato consultas"

    Public Shared Function putQuotes(ByVal string2Quote As String) As String

        Return ControlChars.Quote + string2Quote + ControlChars.Quote

    End Function

    Public Shared Function getDataBaseRef(ByVal tableName As String, ByVal Sociedad As eSociedad) As String

        Dim SociedadNombre As String = NOMBRESOCIEDAD(Sociedad)

        Select Case ConfigurationManager.AppSettings.Get("DBType").ToString

            Case DBTypeHANA
                Return putQuotes(ConfigurationManager.AppSettings.Get("bd_" & NOMBRESOCIEDAD(Sociedad)).ToString) + "." + putQuotes(tableName)

            Case Else
                Return putQuotes(ConfigurationManager.AppSettings.Get("bd_" & NOMBRESOCIEDAD(Sociedad)).ToString) + ".dbo." + putQuotes(tableName)

        End Select

    End Function

    Public Shared Function getDefaultDate() As String

        Select Case ConfigurationManager.AppSettings.Get("DBType").ToString

            Case DBTypeHANA
                Return "CURRENT_DATE"

            Case Else
                Return "GETDATE()"

        End Select

    End Function

    Public Shared Function getDefaultDateWithoutTime() As String

        Select Case ConfigurationManager.AppSettings.Get("DBType").ToString

            Case DBTypeHANA
                Return "CURRENT_DATE"

            Case Else
                Return "Cast(GETDATE() as date)"

        End Select

    End Function

    Public Shared Function getDateAsString_yyyyMMdd(ByVal tableNPto As String, ByVal field As String) As String

        Select Case ConfigurationManager.AppSettings.Get("DBType").ToString

            Case DBTypeHANA
                Return "TO_NVARCHAR(" & tableNPto & putQuotes(field) & ",'YYYYMMDD')"

            Case Else
                Return "CONVERT(NVARCHAR, " & tableNPto & putQuotes(field) & ", 112)"

        End Select

    End Function

    Public Shared Function getStringAsNumber(ByVal tableNPto As String, ByVal field As String) As String

        Select Case ConfigurationManager.AppSettings.Get("DBType").ToString

            Case DBTypeHANA
                Return "MAX(TO_NUMBER(" & tableNPto & putQuotes(field) & "))"

            Case Else
                Return "MAX(Cast(" & tableNPto & putQuotes(field) & " as numeric))"

        End Select

    End Function

    Public Shared Function setNumberAsString(ByVal tableNPto As String, ByVal field As String) As String

        Select Case ConfigurationManager.AppSettings.Get("DBType").ToString

            Case DBTypeHANA
                Return "TO_NVARCHAR(" & tableNPto & putQuotes(field) & ")"

            Case Else
                Return "Cast(" & tableNPto & putQuotes(field) & " as nvarchar)"

        End Select

    End Function

    Public Shared Function getWithNoLock() As String

        Select Case ConfigurationManager.AppSettings.Get("DBType").ToString

            Case DBTypeHANA
                Return ""

            Case Else
                Return "WITH(NOLOCK)"

        End Select

    End Function

#End Region

End Class
