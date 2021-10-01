SELECT customer_info.cid AS cid, NAME, lid, email
FROM customer_info
INNER JOIN customer_onlineinfo
ON customer_info.cid = customer_onlineinfo.cid
WHERE email IS NOT NULL;