CREATE VIEW "BASEDATOS_HANA"."SEI_VIEW_DW_ACTIVOS_FIJOS_GRUPOS" AS
SELECT
COALESCE(T0."Code",'') As ID,
COALESCE(T0."Descr",'') As NOMBRE
FROM ODTP T0
WHERE 1=1 


