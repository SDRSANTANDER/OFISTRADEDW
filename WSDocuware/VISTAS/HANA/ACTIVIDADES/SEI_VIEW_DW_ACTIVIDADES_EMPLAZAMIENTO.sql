CREATE VIEW "BASEDATOS_HANA"."SEI_VIEW_DW_ACTIVIDADES_EMPLAZAMIENTO" AS 
SELECT
COALESCE(T0."Code", -1) As ID,
COALESCE(T0."Name",'') As DESCRIPCION
FROM OCLO T0 
WHERE 1=1 
And T0."Locked"='Y'

