CREATE TEMPORARY TABLE temp_table AS (
    SELECT 
        variable1, 
        variable2, 
        variable3,
        COUNT(*) AS Count,
        SUM(CASE WHEN variable2 = variable3 THEN 1 ELSE 0 END) AS new_column_name
    FROM 
        imported_table
    GROUP BY 
        variable1, variable2, variable3
    );
