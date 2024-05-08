-- Window Function for colum1-column2 (>70%) KPI Calculation
SELECT 
	jsi.*,
	CASE
	 	WHEN jsi.column1= jsi.column2 THEN
	 	100* (SELECT
	 		COUNT(*) * 1.0/ 
	 		(SELECT
	 		COUNT(*)
	 		FROM table_month_name)
	 	FROM (
	 		SELECT
	 			mainquery.column1,
	 			COUNT(*) * 1.0 / 
	 				(SELECT COUNT(*)
	 				FROM table_month_name AS 
	subquery
	 				WHERE mainquery.column1 = subquery.column1
	 			) AS Avr_frequency
	 		FROM table_month_name AS mainquery
	 		WHERE mainquery.column1 = mainquery.column2
	 		GROUP BY mainquery.column1
	 		HAVING COUNT(*) * 1.0 / 
	 		(SELECT
	 		COUNT(*)
	 		FROM table_month_name AS subquery 
	 		WHERE mainquery.column1= subquery.column1) > 0.7
	 	) AS sub)
	 ELSE NULL 
	END AS new_column_name
FROM table_month_name jsi;
/*ORDER BY column1, column2*/
