CREATE VIEW "BASEDATOS_HANA"."SEI_VIEW_DW_ACTIVIDADES_ASUNTO" AS 
SELECT
COALESCE(T0."Code", -1) As ID,
COALESCE(T0."Name",'') As DESCRIPCION,
COALESCE(T0."Type",-1) As TIPO
FROM OCLS T0
WHERE 1=1 
And T0."Active"='Y'
