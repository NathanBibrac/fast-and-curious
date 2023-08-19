SELECT DISTINCT a.question_id,a.answer_text || ' ou ' || b.answer_text || ' ?' as question_text 
FROM questions a JOIN questions b ON a.question_id = b.question_id
WHERE a.answer_order = 1
AND b.answer_order = 2