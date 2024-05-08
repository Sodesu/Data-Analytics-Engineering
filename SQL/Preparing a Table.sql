CREATE TABLE table_name (column1 text, column2 text);

INSERT INTO table_name (column1, column2)
SELECT "column1", "column2" from import_table;

ALTER TABLE table_name RENAME TO new_table_name;

DELETE FROM new_table_name
WHERE column1 = '' OR column1 = ''

SELECT *
FROM new_table_name;
