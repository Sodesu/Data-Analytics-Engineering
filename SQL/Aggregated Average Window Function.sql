SELECT 
    column1, 
    column2, 
    CASE 
        WHEN column1 = column2 THEN
        AVG(CASE WHEN column1 = column2) THEN 1.0 ELSE 0 END) OVER ()
    ELSE NULL
    END AS new_column_name
FROM 
    renamed_temp_table
GROUP BY 
    new_column_name,
    column1, 
    column2
);
