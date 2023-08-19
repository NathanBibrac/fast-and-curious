WITH questions_text AS(

SELECT DISTINCT a.question_id,a.answer_text || ' ou ' || b.answer_text || ' ?' as question_text 
FROM questions a JOIN questions b ON a.question_id = b.question_id
WHERE a.answer_order = 1
AND b.answer_order = 2

)

SELECT a.answer_id,a.user_id,q.question_id,a.answer,u.username,qt.question_text,q.answer_text,u.is_active 
FROM answers a JOIN questions q ON a.answer = q.answer_order AND a.question_id = q.question_id
               JOIN users u ON a.user_id = u.user_id
               JOIN questions_text qt ON a.question_id = qt.question_id
ORDER BY a.user_id,a.question_id
