import sqlite3
from flask import Flask, render_template, request, redirect
from datetime import datetime

DATABASE = 'D:\\Lab\\Python\\web\\fnc\\data\\fast_and_curious.db'
app = Flask(__name__)

def ts_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def is_username_taken(username):

    print(f'{ts_str()} - Checking if username ({username}) is taken')

    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()
    cursor.execute("SELECT username FROM users WHERE username = ?", (username,))

    result = cursor.fetchone()
    conn.close()

    if(result):
        print(f'{ts_str()} - Username ({username}) is taken')
        return True
    else:
        print(f'{ts_str()} - Username ({username}) is not taken')
        return False

def add_new_username(username):

    print(f'{ts_str()} - Setting new active player to ({username})')

    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()
    
    if(is_username_taken(username)):
        return False

    cursor.execute("UPDATE users SET is_active=0 WHERE is_active=1")
    cursor.execute("INSERT INTO users (username,is_active) VALUES (?,1)", (username,))
    cursor.execute("SELECT username FROM users WHERE is_active=1")
    result = cursor.fetchone()

    if(result):
        conn.commit()
        print(f'{ts_str()} - New active player set to ({username})')
    else:
        print(result)
        conn.rollback()
        print(f'{ts_str()} - Failed to set new active player to ({username})')
    conn.close()

    return True

def get_active_username():
    
    print(f'{ts_str()} - Getting active player')

    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()
    cursor.execute("SELECT username FROM users WHERE is_active=1")
    result = cursor.fetchone()
    conn.close()

    if(result):
        print(f'{ts_str()} - Active player is ({result[0]})')
        return result[0]
    else:
        print(f'{ts_str()} - No active player')
        return None

def get_question_id_user_id(username):

    print(f'{ts_str()} - Getting question id and user id for ({username})')

    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()
    cursor.execute("SELECT MAX(question_id) FROM answers WHERE user_id = (SELECT user_id FROM users WHERE username = ?)", (username,))
    result = cursor.fetchone()
    conn.close()
    print(result[0])
    if(result[0] is not None):
        print(f'{ts_str()} - Got question id ({result[0]}) for ({username})')
        return (result[0]+1)
    else:
        print(f'{ts_str()} - No question id and user id for ({username})')
        return (0)

def check_next_question(i):

    print(f'{ts_str()} - Checking if question ({i}) exists')
    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()
    cursor.execute("SELECT question_id FROM questions WHERE question_id = ?", (i,))
    question_id = cursor.fetchone()
    conn.close()

    if(question_id[0]<10):
        print(f'{ts_str()} - Question ({i}) exists')
        return True
    else:
        print(f'{ts_str()} - Question ({i}) does not exist')
        return False


def get_questions(i):

    print(f'{ts_str()} - Getting questions')

    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()
    cursor.execute("SELECT answer_text, answer_order FROM questions WHERE question_id = ?", (i,))
    result = cursor.fetchall()
    conn.close()

    if(result):
        print(f'{ts_str()} - Got {len(result)} questions')
        print(result)
        print(f'{result[0][0]} ou {result[1][0]} ?')
        return result
    else:
        print(f'{ts_str()} - No questions')
        return None


def add_answer(username, question_id, answer):

    print(f'{ts_str()} - Adding answer ({answer}) for ({username}) to question ({question_id})')

    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()
    cursor.execute("SELECT user_id FROM users WHERE username = ?", (username,))
    user_id = cursor.fetchone()[0]
    cursor.execute("INSERT INTO answers (user_id, question_id, answer) VALUES (?, ?, ?)", (user_id, question_id, answer))
    conn.commit()
    conn.close()

    print(f'{ts_str()} - Answer added')

    return True

def get_answers(username):
    
        print(f'{ts_str()} - Getting answers for ({username})')
    
        conn = sqlite3.connect(DATABASE)
        cursor = conn.cursor()
        query = """
        
        SELECT a.answer_text || ' ou ' || b.answer_text as question_text,
        CASE ans.answer
            WHEN 1 THEN a.answer_text
            WHEN 2 THEN b.answer_text
        END AS answer_text
             
        FROM answers ans JOIN questions a ON a.question_id = ans.question_id 
                         JOIN questions b ON b.question_id = ans.question_id 
        
        WHERE a.answer_order = 1 AND b.answer_order = 2 
        AND ans.user_id = (SELECT user_id FROM users WHERE username = ?)
        
        """
        cursor.execute(query, (username,))
        result = cursor.fetchall()
        conn.close()
    
        if(result):
            print(f'{ts_str()} - Got {len(result)} answers')
            return result
        else:
            print(f'{ts_str()} - No answers')
            return None


@app.route('/')
def index():
    return render_template('index.html')

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        if add_new_username(username):
            return redirect('/question')
        else:
            #return render_template('login.html', error='Username already taken.')
            answer_dict = get_answers(username)
            rendered_lines = []
            for answer in answer_dict:
                rendered_lines.append(f'{answer[0]} : {answer[1]}')
            return render_template('thankyou.html', answers=rendered_lines)

    return render_template('login.html')

@app.route('/question', methods=['GET', 'POST'])
def question():
    if request.method == 'POST':
        username = get_active_username()
        question_id = get_question_id_user_id(username)
        answer = request.form['answer']
        ### Partie à modifier en fonction du template. 
        #Si le premier bullet est selectionné alors awser = 1 sinon answer = 2
        if answer == '1':
            answer = 1
        else:
            answer = 2
        add_answer(username, question_id, answer)

        if check_next_question(question_id+1):
            return redirect('/question')
        else:
            return redirect('/thankyou')
    else :
        print(get_active_username())
        print(get_question_id_user_id(get_active_username()))
        answer1 = get_questions(get_question_id_user_id(get_active_username()))[0][0]
        answer2 = get_questions(get_question_id_user_id(get_active_username()))[1][0]
        question = answer1 + ' ou ' + answer2 + ' ?'
        get_questions(get_question_id_user_id(get_active_username()))
        return render_template('question.html', question=question, answer1=answer1, answer2=answer2)
    
@app.route('/thankyou')

def thankyou( methods=['GET', 'POST']):
    if request.method == 'POST':
        return redirect('/login')
    
    answer_dict = get_answers(get_active_username())
    rendered_lines = []
    for answer in answer_dict:
        rendered_lines.append(f'{answer[0]} : {answer[1]}')
    return render_template('thankyou.html', answers=rendered_lines)


def main():
    app.run(debug=True)

if __name__ == '__main__':
    main()