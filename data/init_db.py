import sqlite3
import pandas as pd
from datetime import datetime

## Variables d'environnement

#sqlite database path

DB = r'D:\Lab\Python\web\fnc\data\fast_and_curious.db'
EXCEL = r'D:\Lab\Python\web\fnc\data\xl_fnf_upd.xlsx'

## Functions

#timestamp string
def ts_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# Create users, questions, and answers tables


def dump_exisiting_tables():
    
        print(f'{ts_str()}: Dumping existing tables...')
    
        conn = sqlite3.connect(DB)
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
        tables = cursor.fetchall()
        for table in tables:
            print(f'{table[0]}')
            cursor.execute(f"SELECT * FROM {table[0]}")
            rows = cursor.fetchall()
            for row in rows:
                print(row)
        conn.close()
    
        print(f'{ts_str()}: Existing tables dumped.')

def create_tables(): 

    print(f'{ts_str()}: Creating database tables...')

    conn = sqlite3.connect(DB)
    cursor = conn.cursor()
    cursor.execute('''CREATE TABLE IF NOT EXISTS questions (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        question_id INTEGER,
                        answer_text TEXT,
                        answer_order INTEGER
                    )''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS users (
                        user_id INTEGER PRIMARY KEY AUTOINCREMENT,
                        username TEXT,
                        is_active INTEGER DEFAULT 1
                    )''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS answers (
                        answer_id INTEGER PRIMARY KEY AUTOINCREMENT,
                        user_id INTEGER,
                        question_id INTEGER,
                        answer TEXT
                    )''')
    
    cursor.execute("INSERT OR REPLACE INTO users (username,is_active) VALUES ('admin',1)")

    conn.commit()
    conn.close()

    print(f'{ts_str()}: Database tables created.')


def insert_questions():
    
    print(f'{ts_str()}: Inserting questions from {EXCEL}...')

    data_excel = pd.read_excel(EXCEL)
    questions = []
    
    for index, row in data_excel.iterrows():

        question_id = index
        first_text = row['First']
        second_text = row['Second']
    
        questions.append([question_id,first_text,1])
        questions.append([question_id,second_text,2])
    
    conn = sqlite3.connect(DB)
    cursor = conn.cursor()
    cursor.executemany("INSERT INTO questions (question_id, answer_text, answer_order) VALUES (?, ?, ?)", questions)
    conn.commit()
    conn.close()
    
    print(f'{ts_str()}: Questions inserted in database.')

    # Display database tables and column list

def display_table_col():

    print(f'ts_str(): Displaying database tables and columns...')

    conn = sqlite3.connect(DB)
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
    tables = cursor.fetchall()
    for table in tables:
        print(f'{table[0]}')
        cursor.execute(f"SELECT * FROM {table[0]}")
        columns = [description[0] for description in cursor.description]
        print(columns)
    conn.close()

    print(f'{ts_str()}: Database tables and columns displayed.')



## Main

def main():
    print(f'{ts_str()}: Starting init_db.py')
    create_tables()
    insert_questions()
    display_table_col()
    print(f'{ts_str()}: Done Running.')

if __name__ == "__main__":
    main()