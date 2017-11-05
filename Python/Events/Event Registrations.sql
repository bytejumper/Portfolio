/* ****************************************************************************
** Event Registrations
**
** Determines list of event registrants who are affiliated with Engineering
**	through degree, giving, or employment
** 		- @eventLookupID
** 			LookupID of applicable event
**
**	* Use in conjunction with StandardOuput to display Constituent details	*
**************************************************************************** */
USE BBCRM_Prod_RPT_BBDW;

DECLARE @eventLookupID varchar(10)
Set @eventLookupID = '';

WITH cte AS (
SELECT DISTINCT er.CONSTITUENTDIMID
FROM BBDW.DIM_EVENT AS e
	JOIN BBDW.FACT_EVENTREGISTRANT AS er
		ON er.EVENTDIMID = e.EVENTDIMID
	JOIN BBDW.DIM_EVENTREGISTRATION AS reg
		ON reg.EVENTREGISTRATIONDIMID = er.EVENTREGISTRATIONDIMID
	JOIN BBDW.V_FACT_AFFILIATIONRATINGBYADVANCEMENTUNIT_EXT AS a
		ON a.CONSTITUENTDIMID = er.CONSTITUENTDIMID
WHERE e.EVENTLOOKUPID = @eventLookupID
	AND reg.WILLNOTATTEND = 0
	AND a.ADVANCEMENTUNITCODE = 'U ENGR'
	AND (a.EMPLOYMENT > 0 OR a.GIVING > 0 OR a.DEGREE > 0)
)

SELECT s.*
FROM cte
	CROSS APPLY UIUC..StandardOuput(cte.CONSTITUENTDIMID) AS s
