CREATE VIEW "BASEDATOS_HANA"."SEI_VIEW_DW_PORTES" AS 
SELECT
COALESCE(T0."ExpnsCode",-1) As ID,
COALESCE(T0."ExpnsName", '') As NOMBRE
FROM OEXD T0 
WHERE 1=1 

