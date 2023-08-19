DELETE FROM answers 
WHERE user_id IN (
      SELECT user_id 
      FROM USERS 
      WHERE username IN ('test','test6','test7','chaussette','Ok')
      )