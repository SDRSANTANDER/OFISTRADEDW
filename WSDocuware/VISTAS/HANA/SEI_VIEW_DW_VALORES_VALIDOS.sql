CREATE VIEW "BASEDATOS_HANA"."SEI_VIEW_DW_VALORES_VALIDOS" AS
SELECT
COALESCE(T0."TableID",'') As TABLAID,
COALESCE(T0."FieldID",-1) As CAMPOID,
COALESCE(T0."AliasID",'') As ALIASID,
COALESCE(T0."Descr",'') As NOMBRE,
COALESCE(T0."Dflt",'') As VALORVALIDODEFECTO,
(SELECT STRING_AGG(COALESCE(T."DESCRIPCION",''),'||')
    FROM (
        SELECT
        COALESCE(X0."FldValue",'') AS DESCRIPCION
        FROM UFD1 X0
        WHERE 1=1
        And X0."TableID"=T0."TableID"
        And X0."FieldID"=T0."FieldID"
        GROUP BY X0."FldValue"
    ) As T
) As VALORESVALIDOSID,
(SELECT
    STRING_AGG(COALESCE(T."DESCRIPCION",''),'||')
    FROM (
        SELECT
        COALESCE(X0."Descr",'') AS DESCRIPCION
        FROM UFD1 X0
        WHERE 1=1
        And X0."TableID"=T0."TableID"
        And X0."FieldID"=T0."FieldID"
        GROUP BY X0."Descr"
      ) As T
) As VALORESVALIDOSNOMBRE
FROM CUFD T0
JOIN UFD1 T1 ON T1."TableID"=T0."TableID" And T1."FieldID"=T0."FieldID"
WHERE 1=1
GROUP BY
T0."TableID",
T0."FieldID",
T0."AliasID",
T0."Descr",
T0."Dflt"

