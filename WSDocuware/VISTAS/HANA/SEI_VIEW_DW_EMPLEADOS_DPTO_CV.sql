CREATE VIEW "BASEDATOS_HANA"."SEI_VIEW_DW_EMPLEADOS_DPTO_CV" AS 
SELECT
COALESCE(T0."SlpCode",0) As ID,
COALESCE(T0."SlpName",'') As NOMBRE,
COALESCE(T0."EmpID",COALESCE(T1."empID",0)) As EMPID,
CASE 
	WHEN COALESCE(T0."Telephone",'')='' THEN COALESCE(T0."Mobil",'') 
	ELSE COALESCE(T0."Telephone",'') 
END As TELEFONO,
COALESCE(T0."Email",'') As CORREOE
FROM OSLP T0
LEFT JOIN OHEM T1 ON T1."salesPrson"=T0."SlpCode"
WHERE 1=1 
And T0."Active"='Y'


--SELECT
--COALESCE(T0."SlpCode",0) As ID,
--COALESCE(T0."SlpName",'') As NOMBRE,
--COALESCE(T0."EmpID",0) As EMPID,
--CASE 
--	WHEN COALESCE(T0."Telephone",'')='' THEN COALESCE(T0."Mobil",'') 
--	ELSE COALESCE(T0."Telephone",'') 
--END As TELEFONO,
--COALESCE(T0."Email",'') As CORREOE
--FROM OSLP T0 
--WHERE 1=1 
--And T0."Active"='Y'