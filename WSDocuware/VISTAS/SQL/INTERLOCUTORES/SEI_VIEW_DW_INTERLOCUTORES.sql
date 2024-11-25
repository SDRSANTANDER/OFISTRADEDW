CREATE VIEW SEI_VIEW_DW_INTERLOCUTORES AS 
SELECT
COALESCE(T0."CardCode", '') As ID,
COALESCE(T0."CardType", '') As TIPO,
COALESCE(T0."CardName", '') As RAZONSOCIAL,
COALESCE(T0."CardFName", '') As NOMBREEXTRANJERO,
COALESCE(T0."LicTradNum", '') As NIF,
COALESCE(T0."U_SEIICDW", '') As ICDW,
COALESCE(T0."LangCode", 0) As IDIOMA,
COALESCE(T0."SlpCode", -1) As ENCARGADO,
COALESCE(T1."email", '') As CORREOE,
COALESCE(T2."City", '') As FACTURACIUDAD,
COALESCE(T0."ECVatGroup", '') As IVAGRUPO,
COALESCE(T4."Rate", 0) As IVAPORCENTAJE,
COALESCE(T3."PymCode", '') As VIAPAGO,
COALESCE(T0."GroupNum", -1) As CONDICIONPAGO,
COALESCE(T0."U_SEIRecDW", 'N') As RECURRENTE,
COALESCE(Cast(T0."AliasName" as nvarchar(max)), '') As AUXILIAR,
COALESCE(T0."U_SEIImpDW", 0) As IMPORTE,
CASE WHEN COALESCE(T0."frozenFor",'N')='Y' THEN 'N' ELSE 'S' END As ACTIVO
FROM OCRD T0 WITH(NOLOCK)
LEFT JOIN OHEM T1 WITH(NOLOCK) ON T0."SlpCode" = T1."salesPrson" 
LEFT JOIN CRD1 T2 WITH(NOLOCK) ON T2."CardCode" = T0."CardCode" And T2."Address"=T0."BillToDef"
LEFT JOIN CRD2 T3 WITH(NOLOCK) ON T3."CardCode" = T0."CardCode" And T3."PymCode"=T0."PymCode"
LEFT JOIN OVTG T4 WITH(NOLOCK) ON T4."Code"=T0."ECVatGroup"
WHERE 1=1 
And T0."frozenFor"<>'Y'
--And COALESCE(T0."ValidFrom",GETDATE())<=GETDATE()
--And COALESCE(T0."ValidTo",GETDATE())>=GETDATE()
GROUP BY 
T0."CardCode",
T0."CardType",
T0."CardName",
T0."CardFName",
T0."LicTradNum",
T0."U_SEIICDW",
T0."LangCode",
T0."SlpCode",
T1."email",
T2."City",
T3."PymCode",
T0."GroupNum",
T0."ECVatGroup",
T4."Rate",
Cast(T0."AliasName" as nvarchar(max)),
T0."U_SEIRecDW",
T0."U_SEIImpDW",
T0."frozenFor"

