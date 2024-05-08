--column1-column2(>70%) KPI Calculation
SELECT 100 * COUNT(*) * 1.0 /(SELECT COUNT(*) FROM table_month_name) as new_column_name
FROM(
	SELECT 
		column1, 
		column2,
		COUNT(*) * 1.0 / (
			SELECT COUNT(*) 
			FROM new_column_name AS subquery 
			WHERE mainquery.column1= subquery.column1
		) AS Avr_freq
	FROM table_month_name AS mainquery
	WHERE column1 = column2
	GROUP BY column2 -- Since the condition is set where the values from both columns are equal, the group by can be either column
) AS sub
WHERE Avr_freq > 0.7;  /*Avr_freq < 0.7  100 - (COUNT(*) * 1.0 / (SELECT COUNT(*) FROM table_month_name) * 100);*/
