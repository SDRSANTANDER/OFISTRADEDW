CREATE VIEW "BASEDATOS_HANA"."SEI_VIEW_DW_LLAMADAS_ORIGEN" AS 
SELECT
COALESCE(T0."originID", -1) As ID,
COALESCE(T0."Name",'') As NOMBRE,
COALESCE(T0."Descriptio",'') As DESCRIPCION
FROM OSCO T0
WHERE 1=1 
And T0."Active"='Y'

