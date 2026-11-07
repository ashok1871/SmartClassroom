import os, hmac, hashlib, secrets, sqlite3, requests, html
from pathlib import Path
from datetime import datetime
import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

SCHOOL = "PM SHRI KENDRIYA VIDYALAYA NO. 2 ARMAPUR, KANPUR"
APP_TITLE = "Smart Classroom Teaching–Learning Portal"
BASE = Path(__file__).parent
DB_PATH = Path(os.getenv("CLASSROOM_DB_PATH", str(BASE / "data" / "classroom.db")))
DB_PATH.parent.mkdir(parents=True, exist_ok=True)

st.set_page_config(page_title=SCHOOL, page_icon="🎓", layout="wide")

I18N = {
"English": {
"portal":"Smart Classroom Teaching–Learning Portal","login":"Sign in","username":"Username","password":"Password",
"sign_in":"Sign in","logout":"Log out","language":"Language / भाषा","role":"Role","teacher":"Teacher",
"student":"Student","dashboard":"Dashboard","resources":"Learning Resources","quiz":"Online Quiz",
"doubts":"Doubt Corner + AI Assistant","analytics":"Student-wise Analytics","paper":"CBSE Question Paper Generator",
"users":"User Management","settings":"Account & System","inclusive":"Inclusive Learning","welcome":"Welcome",
"add_resource":"Add learning resource","save":"Save","subject":"Subject","class":"Class","topic":"Topic / Chapter",
"title":"Title","type":"Resource type","url":"Resource URL","notes":"Description / Notes",
"add_question":"Add question to question bank","question":"Question","qtype":"Question type","marks":"Marks",
"difficulty":"Difficulty","a":"Option A","b":"Option B","c":"Option C","d":"Option D","answer":"Correct option",
"explanation":"Explanation / marking points","submit":"Submit","choose_quiz":"Choose quiz","attempt":"Submit answers",
"your_score":"Your score","submit_doubt":"Submit doubt","response":"Teacher response","status":"Status",
"generate":"Generate question paper","download":"Download editable DOCX","name":"Full name",
"create_user":"Create user","new_password":"Temporary password","create":"Create account",
"change_password":"Change password","old_password":"Current password","new_password2":"New password",
"confirm_password":"Confirm new password","update":"Update password","ai":"Ask AI learning assistant",
"ai_note":"AI responses are drafts. Verify accuracy and CBSE alignment before use.",
"no_data":"No records yet.","first_admin":"Initial administrator setup","setup_help":"Create the first teacher account. This setup appears only while no account exists.",
"school":"School","open_doubts":"Open doubts","attempts":"Quiz responses","resources_count":"Resources",
"questions_count":"Questions","student_name":"Student name","all_classes":"All classes","all_subjects":"All subjects",
"topic_filter":"Topic filter","paper_title":"Examination / Practice Test","duration":"Time allowed",
"max_marks":"Maximum marks","instructions":"General instructions","add":"Add","delete":"Delete",
"account_created":"Account created successfully.","bad_login":"Incorrect username or password.",
"password_rules":"Use at least 10 characters, including uppercase, lowercase and a number.",
"network":"Deployment and security","no_ai":"AI service is not configured. Set AI_API_KEY in the deployment environment to enable it.",
"ai_failed":"AI assistant could not respond. Check the AI endpoint/key and try again.",
"student_code":"Student ID / roll number","student_only":"This area is available to students after sign-in.",
"teacher_only":"This area is available to teachers after sign-in."
},
"Hindi": {
"portal":"स्मार्ट कक्षा शिक्षण-अधिगम पोर्टल","login":"लॉगिन","username":"यूज़रनेम","password":"पासवर्ड",
"sign_in":"लॉगिन करें","logout":"लॉगआउट","language":"Language / भाषा","role":"भूमिका","teacher":"शिक्षक",
"student":"विद्यार्थी","dashboard":"डैशबोर्ड","resources":"अधिगम संसाधन","quiz":"ऑनलाइन क्विज़",
"doubts":"प्रश्न मंच + AI सहायक","analytics":"विद्यार्थी-वार विश्लेषण","paper":"CBSE प्रश्नपत्र निर्माता",
"users":"यूज़र प्रबंधन","settings":"खाता एवं सिस्टम","inclusive":"समावेशी अधिगम","welcome":"स्वागत है",
"add_resource":"अधिगम संसाधन जोड़ें","save":"सहेजें","subject":"विषय","class":"कक्षा","topic":"विषय-वस्तु / अध्याय",
"title":"शीर्षक","type":"संसाधन प्रकार","url":"संसाधन URL","notes":"विवरण / टिप्पणियाँ",
"add_question":"प्रश्न बैंक में प्रश्न जोड़ें","question":"प्रश्न","qtype":"प्रश्न का प्रकार","marks":"अंक",
"difficulty":"कठिनाई स्तर","a":"विकल्प A","b":"विकल्प B","c":"विकल्प C","d":"विकल्प D","answer":"सही विकल्प",
"explanation":"व्याख्या / मूल्यांकन बिंदु","submit":"जमा करें","choose_quiz":"क्विज़ चुनें","attempt":"उत्तर जमा करें",
"your_score":"आपका स्कोर","submit_doubt":"प्रश्न जमा करें","response":"शिक्षक का उत्तर","status":"स्थिति",
"generate":"प्रश्नपत्र बनाएँ","download":"संपादन योग्य DOCX डाउनलोड करें","name":"पूरा नाम",
"create_user":"यूज़र बनाएँ","new_password":"अस्थायी पासवर्ड","create":"खाता बनाएँ",
"change_password":"पासवर्ड बदलें","old_password":"वर्तमान पासवर्ड","new_password2":"नया पासवर्ड",
"confirm_password":"नए पासवर्ड की पुष्टि करें","update":"पासवर्ड अपडेट करें","ai":"AI अधिगम सहायक से पूछें",
"ai_note":"AI उत्तर प्रारूप मात्र हैं। उपयोग से पहले तथ्य एवं CBSE अनुरूपता जाँचें।",
"no_data":"अभी कोई रिकॉर्ड उपलब्ध नहीं है।","first_admin":"प्रथम शिक्षक खाते की स्थापना","setup_help":"पहला शिक्षक खाता बनाएँ। यह विकल्प केवल तब दिखेगा जब कोई खाता मौजूद नहीं होगा।",
"school":"विद्यालय","open_doubts":"खुले प्रश्न","attempts":"क्विज़ प्रतिक्रियाएँ","resources_count":"संसाधन",
"questions_count":"प्रश्न","student_name":"विद्यार्थी का नाम","all_classes":"सभी कक्षाएँ","all_subjects":"सभी विषय",
"topic_filter":"विषय-वस्तु फ़िल्टर","paper_title":"परीक्षा / अभ्यास परीक्षा","duration":"समय",
"max_marks":"अधिकतम अंक","instructions":"सामान्य निर्देश","add":"जोड़ें","delete":"हटाएँ",
"account_created":"खाता सफलतापूर्वक बनाया गया।","bad_login":"यूज़रनेम या पासवर्ड गलत है।",
"password_rules":"कम-से-कम 10 अक्षर, एक बड़ा अक्षर, एक छोटा अक्षर और एक अंक रखें।",
"network":"डिप्लॉयमेंट एवं सुरक्षा","no_ai":"AI सेवा कॉन्फ़िगर नहीं है। AI_API_KEY सेट करें।",
"ai_failed":"AI सहायक उत्तर नहीं दे पाया। AI endpoint/key जाँचें।",
"student_code":"विद्यार्थी ID / अनुक्रमांक","student_only":"यह सुविधा लॉगिन के बाद विद्यार्थियों के लिए उपलब्ध है।",
"teacher_only":"यह सुविधा लॉगिन के बाद शिक्षकों के लिए उपलब्ध है।"
}}
def T(k): return I18N[st.session_state.get("lang","English")].get(k,k)
def now(): return datetime.now().strftime("%Y-%m-%d %H:%M:%S")
def conn(): return sqlite3.connect(DB_PATH, timeout=20)
def run(sql, params=()):
    with conn() as c: c.execute(sql, params); c.commit()
def df(sql, params=()):
    with conn() as c: return pd.read_sql_query(sql,c,params=params)

def init_db():
    with conn() as c:
        c.executescript("""
        CREATE TABLE IF NOT EXISTS users(
          id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE NOT NULL,
          full_name TEXT NOT NULL, role TEXT NOT NULL CHECK(role IN ('teacher','student')),
          student_code TEXT DEFAULT '', class_name TEXT DEFAULT '', salt BLOB NOT NULL,
          password_hash BLOB NOT NULL, active INTEGER DEFAULT 1, created_at TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS resources(
          id INTEGER PRIMARY KEY AUTOINCREMENT,title TEXT NOT NULL,subject TEXT,class_name TEXT,
          topic TEXT,resource_type TEXT,url TEXT,notes TEXT,created_at TEXT);
        CREATE TABLE IF NOT EXISTS questions(
          id INTEGER PRIMARY KEY AUTOINCREMENT,title TEXT NOT NULL,subject TEXT,class_name TEXT,
          topic TEXT,question TEXT NOT NULL,question_type TEXT DEFAULT 'MCQ',
          option_a TEXT,option_b TEXT,option_c TEXT,option_d TEXT,answer TEXT,
          explanation TEXT,marks INTEGER DEFAULT 1,difficulty TEXT DEFAULT 'Moderate',created_at TEXT);
        CREATE TABLE IF NOT EXISTS attempts(
          id INTEGER PRIMARY KEY AUTOINCREMENT,user_id INTEGER NOT NULL,question_id INTEGER NOT NULL,
          selected TEXT, is_correct INTEGER, attempted_at TEXT,
          FOREIGN KEY(user_id) REFERENCES users(id), FOREIGN KEY(question_id) REFERENCES questions(id));
        CREATE TABLE IF NOT EXISTS doubts(
          id INTEGER PRIMARY KEY AUTOINCREMENT,user_id INTEGER NOT NULL,subject TEXT,topic TEXT,
          doubt TEXT NOT NULL,response TEXT DEFAULT '',status TEXT DEFAULT 'Open',
          created_at TEXT,responded_at TEXT, FOREIGN KEY(user_id) REFERENCES users(id));
        CREATE TABLE IF NOT EXISTS feedback(
          id INTEGER PRIMARY KEY AUTOINCREMENT,user_id INTEGER,topic TEXT,understanding INTEGER,
          comment TEXT,created_at TEXT);
        """)
        c.commit()
init_db()

def seed_demo_content():
    # Small, clearly labelled demonstration bank; teachers may edit/extend it.
    if int(df("SELECT COUNT(*) n FROM resources").iloc[0].n) == 0:
        demo_resources = [
            ("Demo: Python Data Types", "Computer Science", "XI", "Python basics", "Notes", "",
             "Demonstration resource. Replace or supplement with teacher-approved notes.", now()),
            ("Demo: SQL SELECT Statement", "Informatics Practices", "XII", "SQL", "Notes", "",
             "Demonstration resource for SELECT and WHERE queries.", now())
        ]
        with conn() as c:
            c.executemany("""INSERT INTO resources(title,subject,class_name,topic,resource_type,url,notes,created_at)
                VALUES(?,?,?,?,?,?,?,?)""", demo_resources)
            c.commit()
    if int(df("SELECT COUNT(*) n FROM questions").iloc[0].n) == 0:
        demo_questions = [
            ("Demo Python Quiz", "Computer Science", "XI", "Identifiers",
             "Which is a valid Python identifier?", "MCQ", "2value", "value_2", "class", "value-2",
             "B", "An identifier cannot begin with a digit or be a Python keyword.", 1, "Easy", now()),
            ("Demo SQL Quiz", "Informatics Practices", "XII", "SQL SELECT",
             "Which SQL command retrieves records from a table?", "MCQ", "GET", "OPEN", "SELECT", "SHOW TABLE",
             "C", "SELECT retrieves data from one or more tables.", 1, "Easy", now())
        ]
        with conn() as c:
            c.executemany("""INSERT INTO questions(title,subject,class_name,topic,question,question_type,
                option_a,option_b,option_c,option_d,answer,explanation,marks,difficulty,created_at)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", demo_questions)
            c.commit()

seed_demo_content()

def pw_hash(password, salt=None):
    salt = salt or secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode(), salt, 310000)
    return salt, digest
def valid_password(p):
    return len(p)>=10 and any(x.isupper() for x in p) and any(x.islower() for x in p) and any(x.isdigit() for x in p)
def check_password(p, salt, digest):
    return hmac.compare_digest(pw_hash(p,salt)[1], digest)

def school_header():
    st.markdown(f"""
    <div style="text-align:center;border-bottom:3px solid #1f4e79;padding:7px 4px 10px;margin-bottom:12px">
      <div style="font-size:20px;font-weight:800;letter-spacing:.3px">{SCHOOL}</div>
      <div style="font-size:15px;font-weight:600">SMART CLASSROOM TEACHING–LEARNING PORTAL</div>
    </div>""", unsafe_allow_html=True)

if "lang" not in st.session_state: st.session_state.lang="English"
if "user" not in st.session_state: st.session_state.user=None
school_header()
st.title("🎓 " + T("portal"))
st.caption(SCHOOL)

def login_screen():
    users = int(df("SELECT COUNT(*) n FROM users").iloc[0].n)
    if users == 0:
        st.subheader(T("first_admin")); st.info(T("setup_help"))
        with st.form("first_admin"):
            full=st.text_input(T("name")); uname=st.text_input(T("username"))
            p1=st.text_input(T("password"),type="password"); p2=st.text_input(T("confirm_password"),type="password")
            ok=st.form_submit_button(T("create"))
        if ok:
            if not full.strip() or not uname.strip(): st.error("Enter name and username.")
            elif p1 != p2: st.error("Passwords do not match.")
            elif not valid_password(p1): st.error(T("password_rules"))
            else:
                salt,digest=pw_hash(p1)
                run("INSERT INTO users(username,full_name,role,salt,password_hash,created_at) VALUES(?,?,?,?,?,?)",
                    (uname.strip().lower(),full.strip(),"teacher",salt,digest,now()))
                st.success(T("account_created")); st.rerun()
        return
    st.subheader(T("login"))
    with st.form("login"):
        uname=st.text_input(T("username")).strip().lower()
        pwd=st.text_input(T("password"),type="password")
        go=st.form_submit_button(T("sign_in"),use_container_width=True)
    if go:
        u=df("SELECT * FROM users WHERE username=? AND active=1",(uname,))
        if not u.empty and check_password(pwd,bytes(u.iloc[0].salt),bytes(u.iloc[0].password_hash)):
            row=u.iloc[0]
            st.session_state.user={"id":int(row.id),"username":row.username,"full_name":row.full_name,
                                   "role":row.role,"class_name":row.class_name or "XI","student_code":row.student_code or ""}
            st.rerun()
        else: st.error(T("bad_login"))

def docx_header_footer(doc):
    for sec in doc.sections:
        sec.header_distance=Inches(.3); sec.footer_distance=Inches(.3)
        p=sec.header.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
        r=p.add_run(SCHOOL); r.bold=True; r.font.size=Pt(9)
        p2=sec.footer.paragraphs[0]; p2.alignment=WD_ALIGN_PARAGRAPH.CENTER
        rr=p2.add_run("Smart Classroom Teaching–Learning Portal | " + SCHOOL); rr.font.size=Pt(8)

def make_paper(rows, title, cls, subject, duration, maxmarks, instructions):
    d=Document(); sec=d.sections[0]
    sec.top_margin=Inches(.8); sec.bottom_margin=Inches(.7)
    docx_header_footer(d)
    p=d.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run(SCHOOL); r.bold=True; r.font.size=Pt(14)
    p=d.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run(title); r.bold=True; r.font.size=Pt(13)
    p=d.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(f"Class: {cls}    Subject: {subject}    Time: {duration}    Maximum Marks: {maxmarks}").bold=True
    d.add_heading("General Instructions", level=2)
    for line in [x.strip() for x in instructions.splitlines() if x.strip()]:
        d.add_paragraph(line,style="List Number")
    for i,row in enumerate(rows,1):
        p=d.add_paragraph()
        p.paragraph_format.space_before=Pt(8)
        p.add_run(f"Q{i}. {row['question']} ").bold=True
        p.add_run(f"[{int(row['marks'])} mark(s)]")
        if row["question_type"]=="MCQ":
            for letter,col in zip("ABCD",["option_a","option_b","option_c","option_d"]):
                if row[col]: d.add_paragraph(f"({letter}) {row[col]}")
    path=BASE/"generated_question_paper.docx"; d.save(path)
    return path

def ai_answer(prompt, subject, cls, topic):
    key=os.getenv("AI_API_KEY","").strip()
    if not key: return None, T("no_ai")
    endpoint=os.getenv("AI_API_URL","https://api.openai.com/v1/chat/completions")
    model=os.getenv("AI_MODEL","gpt-4o-mini")
    system=("You are a careful school learning assistant supporting a KVS/CBSE classroom. "
            "Answer in student-friendly language, show steps for Computer Science problems, "
            "do not invent syllabus facts, and state uncertainty. Do not request personal data. "
            "Give hints before full solutions when appropriate.")
    body={"model":model,"messages":[{"role":"system","content":system},
          {"role":"user","content":f"Class: {cls}; Subject: {subject}; Topic: {topic}. Student question: {prompt}"}],
          "temperature":0.2}
    try:
        resp=requests.post(endpoint,headers={"Authorization":f"Bearer {key}","Content-Type":"application/json"},
                           json=body,timeout=45)
        resp.raise_for_status()
        return resp.json()["choices"][0]["message"]["content"],None
    except Exception:
        return None,T("ai_failed")

if st.session_state.user is None:
    with st.sidebar:
        st.selectbox(T("language"),["English","Hindi"],key="lang")
        st.caption(SCHOOL)
    login_screen()
    st.stop()

user=st.session_state.user
with st.sidebar:
    st.selectbox(T("language"),["English","Hindi"],key="lang")
    st.markdown(f"**{user['full_name']}**")
    st.caption("Teacher" if user["role"]=="teacher" else f"Student | Class {user['class_name']}")
    if st.button(T("logout"),use_container_width=True):
        st.session_state.user=None; st.rerun()
    if user["role"]=="teacher":
        pages=[T("dashboard"),T("resources"),T("quiz"),T("doubts"),T("analytics"),T("paper"),T("users"),T("inclusive"),T("settings")]
    else:
        pages=[T("dashboard"),T("resources"),T("quiz"),T("doubts"),T("inclusive"),T("settings")]
    page=st.radio("Navigation",pages,label_visibility="collapsed")
    st.caption(SCHOOL)

# map translated labels to stable identifiers
page_key={T(k):k for k in ["dashboard","resources","quiz","doubts","analytics","paper","users","inclusive","settings"]}
selected=page_key.get(page,"dashboard")

if selected=="dashboard":
    st.subheader(T("welcome")+", "+user["full_name"])
    if user["role"]=="teacher":
        metrics=[
            ("resources_count","SELECT COUNT(*) n FROM resources"),
            ("questions_count","SELECT COUNT(*) n FROM questions"),
            ("open_doubts","SELECT COUNT(*) n FROM doubts WHERE status='Open'"),
            ("attempts","SELECT COUNT(*) n FROM attempts")]
        cols=st.columns(4)
        for col,(label,sql) in zip(cols,metrics): col.metric(T(label),int(df(sql).iloc[0].n))
        st.markdown("### Teaching workflow")
        st.markdown("1. Add learning resources and question-bank items. 2. Students attempt quizzes and submit doubts. 3. Review student-wise analytics and topic gaps. 4. Generate editable practice papers.")
    else:
        st.markdown(f"**Class:** {user['class_name']}  |  **Student ID:** {user['student_code'] or 'Not provided'}")
        attempts=df("SELECT COUNT(*) n, COALESCE(SUM(is_correct),0) c FROM attempts WHERE user_id=?",(user["id"],)).iloc[0]
        c1,c2=st.columns(2); c1.metric("Quiz responses",int(attempts.n)); c2.metric("Correct answers",int(attempts.c))
        myd=df("SELECT COUNT(*) n FROM doubts WHERE user_id=?",(user["id"],)).iloc[0].n
        st.metric("Doubts submitted",int(myd))
    st.info("Use the left navigation to open each classroom module.")

elif selected=="resources":
    st.subheader(T("resources"))
    cls=user["class_name"] if user["role"]=="student" else st.selectbox(T("class"),["XI","XII"])
    res=df("SELECT * FROM resources WHERE class_name=? ORDER BY created_at DESC",(cls,))
    if res.empty: st.info(T("no_data"))
    for _,r in res.iterrows():
        with st.expander(f"{r.title} | {r.subject} | {r.topic}"):
            st.write(r.notes or "")
            if r.url: st.link_button("Open resource",r.url)
    if user["role"]=="teacher":
        st.divider(); st.markdown("### "+T("add_resource"))
        with st.form("add_resource",clear_on_submit=True):
            title=st.text_input(T("title")); subject=st.selectbox(T("subject"),["Computer Science","Informatics Practices","Artificial Intelligence","Other"])
            cls2=st.selectbox(T("class"),["XI","XII"]); topic=st.text_input(T("topic"))
            typ=st.selectbox(T("type"),["Notes","Video","Worksheet","Presentation","Website","Other"])
            url=st.text_input(T("url")); notes=st.text_area(T("notes"))
            if st.form_submit_button(T("save")):
                if title.strip():
                    run("INSERT INTO resources(title,subject,class_name,topic,resource_type,url,notes,created_at) VALUES(?,?,?,?,?,?,?,?)",
                        (title,subject,cls2,topic,typ,url,notes,now())); st.success("Saved."); st.rerun()
                else: st.error("Title is required.")

elif selected=="quiz":
    st.subheader(T("quiz"))
    cls=user["class_name"] if user["role"]=="student" else st.selectbox(T("class"),["XI","XII"],key="quizclass")
    quizzes=df("SELECT DISTINCT title FROM questions WHERE class_name=? ORDER BY title",(cls,))
    if not quizzes.empty:
        qtitle=st.selectbox(T("choose_quiz"),quizzes.title.tolist())
        qs=df("SELECT * FROM questions WHERE class_name=? AND title=? ORDER BY id",(cls,qtitle))
        if user["role"]=="student":
            with st.form("take_quiz"):
                answers={}
                for i,r in qs.iterrows():
                    st.markdown(f"**Q{i+1}. {r.question}** ({r.marks} mark)")
                    if r.question_type=="MCQ":
                        answers[int(r.id)]=st.radio("Choose",["A","B","C","D"],
                            format_func=lambda x,row=r: f"{x}. "+str(row[{"A":"option_a","B":"option_b","C":"option_c","D":"option_d"}[x]]),key=f"ans{r.id}")
                    else:
                        answers[int(r.id)]=st.text_area("Your answer",key=f"ans{r.id}")
                submit=st.form_submit_button(T("attempt"))
            if submit:
                score=0; possible=0
                for _,r in qs.iterrows():
                    chosen=answers[int(r.id)]
                    correct=int(r.question_type=="MCQ" and chosen==r.answer)
                    score += correct*int(r.marks); possible += int(r.marks)
                    run("INSERT INTO attempts(user_id,question_id,selected,is_correct,attempted_at) VALUES(?,?,?,?,?)",
                        (user["id"],int(r.id),chosen,correct,now()))
                st.success(f"{T('your_score')}: {score} / {possible} (MCQs auto-graded; descriptive answers require teacher review).")
        else:
            st.dataframe(qs[["question","question_type","topic","marks","difficulty"]],use_container_width=True)
    else: st.info("No questions available for this class.")
    if user["role"]=="teacher":
        st.divider(); st.markdown("### "+T("add_question"))
        with st.form("new_question",clear_on_submit=True):
            title=st.text_input(T("title"),"Practice Quiz")
            subject=st.selectbox(T("subject"),["Computer Science","Informatics Practices","Artificial Intelligence"])
            cls2=st.selectbox(T("class"),["XI","XII"])
            topic=st.text_input(T("topic")); qtype=st.selectbox(T("qtype"),["MCQ","Short Answer","Case-based"])
            question=st.text_area(T("question")); marks=st.number_input(T("marks"),1,10,1)
            diff=st.selectbox(T("difficulty"),["Easy","Moderate","Difficult"])
            oa=ob=oc=od=ans=""
            if qtype=="MCQ":
                oa=st.text_input(T("a")); ob=st.text_input(T("b")); oc=st.text_input(T("c")); od=st.text_input(T("d"))
                ans=st.selectbox(T("answer"),["A","B","C","D"])
            explanation=st.text_area(T("explanation"))
            save=st.form_submit_button(T("save"))
            if save:
                if question.strip() and (qtype!="MCQ" or all(x.strip() for x in [oa,ob,oc,od])):
                    run("""INSERT INTO questions(title,subject,class_name,topic,question,question_type,option_a,option_b,option_c,option_d,answer,explanation,marks,difficulty,created_at)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",(title,subject,cls2,topic,question,qtype,oa,ob,oc,od,ans,explanation,int(marks),diff,now()))
                    st.success("Question added."); st.rerun()
                else: st.error("Enter question and all MCQ options.")

elif selected=="doubts":
    st.subheader(T("doubts"))
    if user["role"]=="student":
        with st.form("new_doubt",clear_on_submit=True):
            subject=st.selectbox(T("subject"),["Computer Science","Informatics Practices","Artificial Intelligence"])
            topic=st.text_input(T("topic")); doubt=st.text_area(T("question"))
            submit=st.form_submit_button(T("submit_doubt"))
        if submit and doubt.strip():
            run("INSERT INTO doubts(user_id,subject,topic,doubt,status,created_at) VALUES(?,?,?,?,?,?)",
                (user["id"],subject,topic,doubt,"Open",now())); st.success("Submitted."); st.rerun()
        mine=df("SELECT * FROM doubts WHERE user_id=? ORDER BY created_at DESC",(user["id"],))
        st.markdown("### My questions")
        for _,r in mine.iterrows():
            with st.expander(f"{r.topic or 'General'} | {r.status}"):
                st.write(r.doubt)
                if r.response: st.success(T("response")+": "+r.response)
        st.divider(); st.markdown("### "+T("ai"))
        st.caption(T("ai_note"))
        with st.form("ai_form"):
            subj=st.selectbox(T("subject"),["Computer Science","Informatics Practices","Artificial Intelligence"],key="aisub")
            topic2=st.text_input(T("topic"),key="aitopic"); prompt=st.text_area(T("question"),key="aiprompt")
            ask=st.form_submit_button(T("ai"))
        if ask and prompt.strip():
            answer,error=ai_answer(prompt,subj,user["class_name"],topic2)
            if error: st.warning(error)
            else: st.markdown(answer); st.caption(T("ai_note"))
    else:
        all_d=df("""SELECT d.*,u.full_name,u.username FROM doubts d JOIN users u ON d.user_id=u.id ORDER BY d.created_at DESC""")
        if all_d.empty: st.info(T("no_data"))
        for _,r in all_d.iterrows():
            with st.expander(f"#{r.id} | {r.full_name} | {r.topic} | {r.status}"):
                st.write(r.doubt); reply=st.text_area(T("response"),value=r.response or "",key=f"reply{r.id}")
                if st.button(T("save"),key=f"replybtn{r.id}"):
                    run("UPDATE doubts SET response=?,status='Answered',responded_at=? WHERE id=?",(reply,now(),int(r.id)))
                    st.success("Response saved."); st.rerun()

elif selected=="analytics":
    st.subheader(T("analytics"))
    attempts=df("""SELECT u.full_name,u.username,u.class_name,q.subject,q.topic,q.title,q.question,
                   a.is_correct,a.attempted_at FROM attempts a JOIN users u ON a.user_id=u.id
                   JOIN questions q ON a.question_id=q.id ORDER BY a.attempted_at DESC""")
    if attempts.empty: st.info(T("no_data"))
    else:
        summary=attempts.groupby(["full_name","username","class_name"]).agg(
            Responses=("is_correct","count"),Correct=("is_correct","sum")).reset_index()
        summary["Accuracy (%)"]=(summary.Correct/summary.Responses*100).round(1)
        st.dataframe(summary,use_container_width=True)
        st.download_button("Download student-wise CSV",summary.to_csv(index=False),"student_analytics.csv","text/csv")
        st.markdown("### Topic-wise learning gaps")
        gaps=attempts.groupby(["subject","topic"]).agg(Responses=("is_correct","count"),Correct=("is_correct","sum")).reset_index()
        gaps["Accuracy (%)"]=(gaps.Correct/gaps.Responses*100).round(1)
        st.dataframe(gaps.sort_values("Accuracy (%)"),use_container_width=True)
        st.bar_chart(gaps.set_index("topic")["Accuracy (%)"])
        st.caption("Use lower-accuracy topics to plan reteaching, peer support and remedial practice.")

elif selected=="paper":
    st.subheader(T("paper"))
    with st.form("paper_builder"):
        title=st.text_input(T("paper_title"),"Class Practice Test")
        cls=st.selectbox(T("class"),["XI","XII"])
        subject=st.selectbox(T("subject"),["Computer Science","Informatics Practices","Artificial Intelligence"])
        topic=st.text_input(T("topic"),"")
        duration=st.text_input(T("duration"),"1 Hour")
        maxmarks=st.number_input(T("max_marks"),1,100,20)
        instructions=st.text_area(T("instructions"),"All questions are compulsory.\nRead each question carefully.\nMarks are indicated against each question.")
        submitted=st.form_submit_button(T("generate"))
    if submitted:
        params=[cls,subject]
        sql="SELECT * FROM questions WHERE class_name=? AND subject=?"
        if topic.strip(): sql+=" AND topic LIKE ?"; params.append("%"+topic.strip()+"%")
        bank=df(sql,tuple(params))
        if bank.empty: st.warning("No matching questions in the question bank.")
        else:
            # Keep within requested maximum; prefer a mix of difficulty levels.
            bank=bank.sort_values(["difficulty","id"])
            chosen=[]; total=0
            for _,row in bank.iterrows():
                m=int(row.marks)
                if total+m<=int(maxmarks): chosen.append(row.to_dict()); total+=m
            if not chosen: st.warning("No question fits the selected maximum marks.")
            else:
                path=make_paper(chosen,title,cls,subject,duration,int(maxmarks),instructions)
                st.success(f"Generated {len(chosen)} questions carrying {total} marks. The generator uses available bank questions and does not guarantee a balanced CBSE blueprint.")
                st.download_button(T("download"),path.read_bytes(),file_name="PM_SHRI_KV2_Armapur_Question_Paper.docx",
                                   mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    st.info("Only teacher-reviewed questions should be added to the question bank. Check chapter coverage, difficulty, marks and CBSE blueprint before printing.")

elif selected=="users":
    st.subheader(T("users"))
    users=df("SELECT id,username,full_name,role,student_code,class_name,active,created_at FROM users ORDER BY role,full_name")
    st.dataframe(users,use_container_width=True)
    st.markdown("### "+T("create_user"))
    with st.form("create_user",clear_on_submit=True):
        full=st.text_input(T("name")); uname=st.text_input(T("username")).strip().lower()
        role=st.selectbox(T("role"),["student","teacher"])
        cls=st.selectbox(T("class"),["XI","XII"])
        code=st.text_input(T("student_code"))
        pwd=st.text_input(T("new_password"),type="password")
        create=st.form_submit_button(T("create"))
    if create:
        if not full.strip() or not uname: st.error("Name and username are required.")
        elif not valid_password(pwd): st.error(T("password_rules"))
        else:
            salt,digest=pw_hash(pwd)
            try:
                run("INSERT INTO users(username,full_name,role,student_code,class_name,salt,password_hash,created_at) VALUES(?,?,?,?,?,?,?,?)",
                    (uname,full.strip(),role,code,cls,salt,digest,now()))
                st.success(T("account_created"))
            except sqlite3.IntegrityError: st.error("Username already exists.")
    st.warning("Share temporary passwords privately. Ask users to change them after first sign-in. Do not share student passwords in public.")

elif selected=="inclusive":
    st.subheader(T("inclusive"))
    size=st.slider("Reading text size / पाठ का आकार",14,32,20)
    contrast=st.selectbox("Display mode / डिस्प्ले",["Standard","Soft cream","High contrast"])
    content=st.text_area("Paste lesson text / पाठ यहाँ लिखें",height=180)
    bg,fg={"Standard":("#ffffff","#111111"),"Soft cream":("#fff8e7","#111111"),"High contrast":("#111111","#ffffff")}[contrast]
    st.markdown(f'<div style="background:{bg};color:{fg};padding:20px;border-radius:8px;font-size:{size}px;line-height:1.7;white-space:pre-wrap">{html.escape(content)}</div>',unsafe_allow_html=True)
    st.caption("Use browser/operating-system read-aloud and screen-reader tools for additional support.")

elif selected=="settings":
    st.subheader(T("settings"))
    st.markdown("### "+T("change_password"))
    with st.form("change_pw"):
        old=st.text_input(T("old_password"),type="password")
        new=st.text_input(T("new_password2"),type="password")
        confirm=st.text_input(T("confirm_password"),type="password")
        update=st.form_submit_button(T("update"))
    if update:
        row=df("SELECT salt,password_hash FROM users WHERE id=?",(user["id"],)).iloc[0]
        if not check_password(old,bytes(row.salt),bytes(row.password_hash)): st.error("Current password is incorrect.")
        elif new!=confirm: st.error("New passwords do not match.")
        elif not valid_password(new): st.error(T("password_rules"))
        else:
            salt,digest=pw_hash(new)
            run("UPDATE users SET salt=?,password_hash=? WHERE id=?",(salt,digest,user["id"]))
            st.success("Password updated.")
    st.divider(); st.markdown("### "+T("network"))
    st.write("Database location:",str(DB_PATH))
    st.write("AI endpoint:",os.getenv("AI_API_URL","OpenAI-compatible chat completions endpoint"))
    st.write("AI model:",os.getenv("AI_MODEL","gpt-4o-mini"))
    st.caption("Configure AI_API_KEY, AI_API_URL and AI_MODEL as environment variables on the server. Never paste API keys into source code.")
    st.warning("For school deployment, use HTTPS behind a trusted reverse proxy, persistent storage, regular backups, network access controls and school approval. The first teacher account is created only on a fresh database.")

st.divider()
st.caption(SCHOOL + " | PM SHRI School Initiative | Educational prototype — protect student data.")
