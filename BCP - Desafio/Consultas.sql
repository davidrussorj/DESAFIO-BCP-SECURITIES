SELECT COUNT(*) AS quantidade_debentures
FROM debentures
WHERE Data = DATE_SUB(CURDATE(), INTERVAL 1 DAY);

SELECT Data, AVG(Duration) AS duration_media
FROM debentures
GROUP BY Data
ORDER BY Data DESC;

SELECT DISTINCT Código
FROM debentures
WHERE Nome = 'VALE S/A';
