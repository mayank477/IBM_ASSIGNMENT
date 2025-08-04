import streamlit as st
import gspread
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from langchain_google_genai import ChatGoogleGenerativeAI
from langchain_core.messages import HumanMessage

# Streamlit UI
st.title("🤖 AI-Powered Customer Support Agent")
user_email = st.text_input("📧 Enter Email")
user_query = st.text_area("💬 Enter Query")

# Gemini API Key input (secure)
GOOGLE_API_KEY = st.secrets["GOOGLE_API_KEY"]
# GOOGLE_API_KEY ='AIzaSyD0yBTO9JWiO1F0BaJ2j4wYhIC2SgHeqUg'     #mayank.mhatre22@it.sce.edu.in
# GOOGLE_API_KEY ='AIzaSyCt1GdsMzr6Mnmx7BPAiCBYIdykZC9cJxo'       #mayankmhatre473@gmail.com

# Initialize Gemini LLM
llm = ChatGoogleGenerativeAI(
    model="models/gemini-1.5-flash-latest",
    temperature=0.2,
    google_api_key=GOOGLE_API_KEY
)

# Google Sheets Auth
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("dogwood-thought-468010-s5-4cbde8425ac1.json", scope)
client = gspread.authorize(creds)
sheet = client.open("CustomerSupportLog").sheet1

# Add Headers
# Ensure headers exist ONCE (no clearing!)
expected_headers = ["Timestamp", "Query", "Department", "Sentiment", "Priority", "Assigned Employee", "Auto Response"]
existing_headers = sheet.row_values(1)

# If headers are missing or incorrect, raise error instead of modifying sheet
if existing_headers != expected_headers:
    st.error("❌ Google Sheet is not correctly set up. Please fix the headers manually and restart the app.")
    st.stop()


# Define departments and employees
employee_groups = {
    "Billing": {"High": ["Tanish"], "Medium": ["Mayank", "Yash"], "Low": ["Leena", "Ketki"]},
    "Technical Support": {"High": ["Tanish"], "Medium": ["Mayank", "Yash"], "Low": ["Leena", "Ketki"]},
    "Sales": {"High": ["Tanish"], "Medium": ["Mayank", "Yash"], "Low": ["Leena", "Ketki"]}
}
employee_emails = {
    "Tanish": "ay19tani@gmail.com", "Mayank": "mayank.mhatre22@it.sce.edu.in", "Yash": "yash.salunkhe22@it.sce.edu.in",
    "Leena": "leenakatkar60@gmail.com", "Ketki": "ketkimane1806@gmail.com",
    "Tanish": "ay19tani@gmail.com", "Mayank": "mayank.mhatre22@it.sce.edu.in", "Yash": "yash.salunkhe22@it.sce.edu.in",
    "Leena": "leenakatkar60@gmail.com", "Ketki": "ketkimane1806@gmail.com",
    "Tanish": "ay19tani@gmail.com", "Mayank": "mayank.mhatre22@it.sce.edu.inn", "Yash": "yash.salunkhe22@it.sce.edu.in",
    "Leena": "leenakatkar60@gmail.com", "Ketki": "ketkimane1806@gmail.com"
}
#workload_tracker = {emp: 0 for dept in employee_groups.values() for prio in dept.values() for emp in prio}

# Initialize workload tracker based on sheet data
workload_tracker = {emp: 0 for emp in employee_emails}

# Count existing assignments from the Google Sheet
records = sheet.get_all_records()
for row in records:
    assigned_emp = row.get("Assigned Employee", "")
    if assigned_emp in workload_tracker:
        workload_tracker[assigned_emp] += 1

# Define core logic
def classify_department(user_query):
    prompt = f"""You are an AI assistant for customer support. Classify this query:
"{user_query}" into Billing, Technical Support, or Sales. Just give the department name."""
    return llm.invoke([HumanMessage(content=prompt)]).content.strip()

def detect_sentiment_and_priority(user_query):
    prompt = f"""Classify the sentiment (Positive, Neutral, Negative) and map to priority:
"{user_query}". Reply in format: Sentiment, Priority"""
    result = llm.invoke([HumanMessage(content=prompt)]).content.strip()
    try:
        sentiment, priority = map(str.strip, result.split(","))
    except:
        sentiment, priority = "Neutral", "Medium"
    return sentiment.capitalize(), priority.capitalize()

def assign_employee(dept, priority):
    options = employee_groups.get(dept, {}).get(priority, [])
    if not options:
        return "No employee available"
    emp = min(options, key=lambda e: workload_tracker[e])
    workload_tracker[emp] += 1
    return emp

def generate_auto_response(query):
    prompt = f"""Generate a short customer service response for:
"{query}" """
    return llm.invoke([HumanMessage(content=prompt)]).content.strip()

def send_email(to_email, subject, message):
    sender_email = "aavishkar20489@gmail.com"
    sender_password = st.secrets["EMAIL_PASSWORD"]
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(message, 'plain'))

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        st.error(f"❌ Email failed: {e}")
        return False

def log_to_sheet(query, dept, sentiment, priority, employee, auto_reply):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    sheet.append_row([timestamp, query, dept, sentiment, priority, employee, auto_reply])

# Process Submission
if st.button("🚀 Submit"):
    if not user_email or not user_query:
        st.warning("Please fill both Email and Query!")
    else:
        dept = classify_department(user_query)
        sentiment, priority = detect_sentiment_and_priority(user_query)
        assigned = assign_employee(dept, priority)
        auto_reply = generate_auto_response(user_query)

        log_to_sheet(user_query, dept, sentiment, priority, assigned, auto_reply)

        # Email to employee
        if assigned in employee_emails:
            send_email(
                employee_emails[assigned],
                f"New Support Ticket Assigned ({priority})",
                f"You've been assigned a ticket:\n\nQuery: {user_query}\nPriority: {priority}\nDepartment: {dept}"
            )
        
        # Email to user
        send_email(
            user_email,
            "Query Received – AI Support",
            f"Hi,\n\nWe received your query:\n\n\"{user_query}\"\n\nOur support team from {dept} will get back shortly.\n\n!"
        )

        # Output
        st.success("✅ We have received your query, please check your Email")
        # st.write(f"**Department:** {dept}")
        # st.write(f"**Sentiment:** {sentiment}")
        # st.write(f"**Priority:** {priority}")
        # st.write(f"**Assigned Employee:** {assigned}")
        # st.info(f"**Auto Response:** {auto_reply}")
