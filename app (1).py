import os
import re
import pandas as pd
from flask import Flask, request, jsonify
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

QUESTION_BANKS = {}

def clean_text(text):
    if not isinstance(text, str):
        return ""
    return re.sub(r'\s+', ' ', text).strip().lower()

def extract_keywords(text):
    words = re.findall(r'\b[a-zA-Z0-9]{4,}\b', clean_text(text))
    stopwords = {'with', 'that', 'this', 'from', 'they', 'have', 'been', 'which', 'their', 'when', 'what', 'where', 'should'}
    return set(w for w in words if w not in stopwords)

def load_question_banks():
    global QUESTION_BANKS
    
    # --- Tab 4: HR Assessment ---
    try:
        hr_df = pd.read_excel('HR_Hiring_Question_Paper_and_Answer_Key.xlsx', sheet_name='HR Evaluation Paper & Key')
        hr_questions = {}
        q_count = 1
        for idx, row in hr_df.iterrows():
            q_id = str(row.iloc[0]) if pd.notna(row.iloc[0]) else ""
            if q_id.startswith('HR-'):
                q_text = str(row.iloc[2]) if pd.notna(row.iloc[2]) else ""
                criteria = str(row.iloc[3]) if pd.notna(row.iloc[3]) else ""
                answer_key = str(row.iloc[4]) if pd.notna(row.iloc[4]) else ""
                max_score = float(row.iloc[5]) if pd.notna(row.iloc[5]) else 10.0
                
                hr_questions[q_count] = {
                    "question": q_text,
                    "answer_key": answer_key,
                    "max_score": max_score,
                    "keywords": extract_keywords(answer_key + " " + criteria)
                }
                q_count += 1
        QUESTION_BANKS[4] = hr_questions
    except Exception as e:
        print(f"Error loading HR Test Excel: {e}")

    # --- Tab 5: Accounting Assessment ---
    try:
        acc_df = pd.read_excel('Accounting_Test_Question_Bank.xlsx', sheet_name='Accounting_Test_Zone5')
        acc_questions = {}
        for idx, row in acc_df.iterrows():
            q_no = int(row['Q.No']) if pd.notna(row['Q.No']) else (idx + 1)
            q_text = str(row['Question']) if pd.notna(row['Question']) else ""
            model_ans = str(row['Model Answer']) if pd.notna(row['Model Answer']) else ""
            grading = str(row['Grading Key Points']) if pd.notna(row['Grading Key Points']) else ""
            max_score = float(row['Max Marks']) if pd.notna(row['Max Marks']) else 10.0
            
            acc_questions[q_no] = {
                "question": q_text,
                "answer_key": model_ans,
                "max_score": max_score,
                "keywords": extract_keywords(model_ans + " " + grading)
            }
        QUESTION_BANKS[5] = acc_questions
    except Exception as e:
        print(f"Error loading Accounting Test Excel: {e}")

    # --- Tab 6: Stock Market Assessment ---
    try:
        sm_xls = pd.ExcelFile('Stock_Market_Examination_Master_Answer_Key.xlsx')
        sm_questions = {}
        q_idx = 1
        for sheet in sm_xls.sheet_names:
            df = pd.read_excel(sm_xls, sheet)
            for idx, row in df.iterrows():
                q_text = str(row['Question']) if pd.notna(row['Question']) else ""
                master_key = str(row['Comprehensive Master Answer Key']) if pd.notna(row['Comprehensive Master Answer Key']) else ""
                max_score = float(row['Marks']) if pd.notna(row['Marks']) else 10.0
                
                sm_questions[q_idx] = {
                    "question": q_text,
                    "answer_key": master_key,
                    "max_score": max_score,
                    "keywords": extract_keywords(master_key)
                }
                q_idx += 1
        QUESTION_BANKS[6] = sm_questions
    except Exception as e:
        print(f"Error loading Stock Market Test Excel: {e}")

    # --- Tab 7: Real Estate Assessment ---
    try:
        re_df = pd.read_excel('Realestate_manager_question_paper_for_hr_.xlsx', sheet_name='Sheet1')
        re_questions = {}
        q_idx = 1
        for idx, row in re_df.iterrows():
            if pd.notna(row.iloc[1]) or pd.notna(row.iloc[2]):
                q_text = str(row.iloc[2]) if pd.notna(row.iloc[2]) else ""
                answer_key = str(row.iloc[3]) if pd.notna(row.iloc[3]) else ""
                max_score = float(row.iloc[4]) if pd.notna(row.iloc[4]) and str(row.iloc[4]).replace('.','',1).isdigit() else 10.0
                
                re_questions[q_idx] = {
                    "question": q_text,
                    "answer_key": answer_key,
                    "max_score": max_score,
                    "keywords": extract_keywords(answer_key)
                }
                q_idx += 1
        QUESTION_BANKS[7] = re_questions
    except Exception as e:
        print(f"Error loading Real Estate Test Excel: {e}")

    # --- Tab 8: Front Office Manager Assessment ---
    try:
        fo_df = pd.read_excel('Front_end_office_manager_HR_Test_.xlsx', sheet_name='Question Paper & Answers')
        fo_questions = {}
        q_num = 1
        for idx, row in fo_df.iterrows():
            q_id = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            # Only match real question rows like "Q1", "Q2", ... (skip the
            # header row "Q. No." which also starts with "Q" and previously
            # caused this whole tab to fail to load).
            if re.fullmatch(r'Q\d+', q_id):
                q_text = str(row.iloc[1]) if pd.notna(row.iloc[1]) else ""
                model_ans = str(row.iloc[2]) if pd.notna(row.iloc[2]) else ""
                try:
                    max_score = float(row.iloc[3]) if pd.notna(row.iloc[3]) else 5.0
                except (ValueError, TypeError):
                    max_score = 5.0

                fo_questions[q_num] = {
                    "question": q_text,
                    "answer_key": model_ans,
                    "max_score": max_score,
                    "keywords": extract_keywords(model_ans)
                }
                q_num += 1
        QUESTION_BANKS[8] = fo_questions
    except Exception as e:
        print(f"Error loading Front Office Manager Test Excel: {e}")

load_question_banks()

def auto_grade_answer(user_answer, key_info):
    user_text = clean_text(user_answer)
    if not user_text:
        return 0.0, "No Answer Provided"

    keywords = key_info.get("keywords", set())
    max_score = key_info.get("max_score", 10.0)
    
    if not keywords:
        return (max_score * 0.5), "Manual Review Required"

    user_words = set(re.findall(r'\b[a-zA-Z0-9]{4,}\b', user_text))
    matched = user_words.intersection(keywords)
    match_ratio = len(matched) / len(keywords) if len(keywords) > 0 else 0

    word_count = len(user_text.split())
    length_multiplier = min(1.0, word_count / 25.0)

    calculated_score = (match_ratio * 0.70 + length_multiplier * 0.30) * max_score
    final_score = round(min(max_score, max(0.0, calculated_score)), 2)

    return final_score, f"Matched {len(matched)} key criteria terms."

@app.route('/api/get-questions', methods=['GET'])
def get_questions():
    questions_response = {}
    for tab_id, q_dict in QUESTION_BANKS.items():
        questions_response[tab_id] = {
            q_no: info["question"] for q_no, info in q_dict.items()
        }
    return jsonify(questions_response)

@app.route('/api/submit-assessment', methods=['POST'])
def submit_assessment():
    data = request.get_json()
    if not data or 'tab_id' not in data or 'answers' not in data:
        return jsonify({"error": "Invalid payload."}), 400

    tab_id = int(data['tab_id'])
    answers = data['answers']

    if tab_id not in QUESTION_BANKS:
        return jsonify({"error": f"Tab {tab_id} answer key not found."}), 400

    q_bank = QUESTION_BANKS[tab_id]
    total_obtained = 0.0
    total_possible = 0.0
    question_breakdown = {}

    for q_no, q_info in q_bank.items():
        user_ans = answers.get(str(q_no), answers.get(q_no, ""))
        score, feedback = auto_grade_answer(user_ans, q_info)
        total_obtained += score
        max_q_score = q_info["max_score"]
        total_possible += max_q_score

        question_breakdown[str(q_no)] = {
            "score_awarded": score,
            "max_score": max_q_score,
            "feedback": feedback,
            "answer_key": q_info.get("answer_key", "")
        }

    percentage = round((total_obtained / total_possible * 100), 2) if total_possible > 0 else 0.0

    return jsonify({
        "status": "success",
        "tab_id": tab_id,
        "total_score": round(total_obtained, 2),
        "max_score": round(total_possible, 2),
        "percentage": percentage,
        "result": "PASSED" if percentage >= 33 else "FAILED",
        "breakdown": question_breakdown
    })

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)




     # ---  pip install pandas openpyxl flask flask-cors ---
     # --- python app.py ---