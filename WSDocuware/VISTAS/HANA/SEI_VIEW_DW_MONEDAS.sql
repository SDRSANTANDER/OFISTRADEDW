CREATE VIEW "BASEDATOS_HANA"."SEI_VIEW_DW_MONEDAS" AS 
SELECT 
COALESCE(T0."CurrCode",'') As ID,
COALESCE(T0."CurrName",'') As NOMBRE,
COALESCE(T0."DocCurrCod", '') As INTERNACIONAL
FROM OCRN T0
WHERE 1=1 

