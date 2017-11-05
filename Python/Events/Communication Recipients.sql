/* ****************************************************************************
** Communication Recipients
**
** Determines list of communication recipients who are affiliated with Engineering
**	through degree, giving, or employment
** 		- @projectNum
** 			Part of source code that describes project number
**			Formatted: <two letter communication type code><six-digit work order number><one letter series code>
**
**	* Use in conjunction with StandardOuput to display Constituent details	*
**************************************************************************** */
USE BBCRM_Prod_RPT_BBDW;

DECLARE @projectNum varchar(10)
Set @projectNum = '';

WITH cte AS (
SELECT DISTINCT mc.CONSTITUENTDIMID
FROM BBDW.DIM_MARKETINGSOURCECODEPART AS sc
	JOIN BBDW.FACT_MARKETINGCONSTITUENT AS mc
		ON mc.MARKETINGSOURCECODEDIMID = sc.MARKETINGSOURCECODEDIMID
	LEFT OUTER JOIN BBDW.V_FACT_AFFILIATIONRATINGBYADVANCEMENTUNIT_EXT AS a
		ON a.CONSTITUENTDIMID = mc.CONSTITUENTDIMID
WHERE sc.PARTCODE = @projectNum
	AND a.ADVANCEMENTUNITCODE = 'U ENGR'
	AND (a.EMPLOYMENT > 0 OR a.GIVING > 0 OR a.DEGREE > 0)
)

SELECT s.*
FROM cte
	CROSS APPLY UIUC..StandardOuput(cte.CONSTITUENTDIMID) AS s
