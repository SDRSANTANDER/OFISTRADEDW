CREATE VIEW SEI_VIEW_DW_TARJETAS_BANCO AS 
SELECT 
COALESCE(T0."CreditCard",-1) As ID,
COALESCE(T0."CardName",'') As NOMBRE,
COALESCE(T0."AcctCode", '') As CUENTA
FROM OCRC T0 WITH(NOLOCK)
WHERE 1=1 
And T0."Locked"<>'Y'

