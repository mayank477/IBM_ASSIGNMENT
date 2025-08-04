import streamlit as st
import gspread
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from langchain_google_genai import ChatGoogleGenerativeAI
from langchain_core.messages import HumanMessage
import json

# Streamlit UI
st.title("🤖 AI-Powered Customer Support Agent")
user_email = st.text_input("📧 Enter Email")
user_query = st.text_area("💬 Enter Query")

# Load API Keys from Streamlit Secrets
GOOGLE_API_KEY = st.secrets["GOOGLE_API_KEY"]

# Initialize Gemini LLM
llm = ChatGoogleGenerativeAI(
    model="models/gemini-1.5-flash-latest",
    temperature=0.2,
    google_api_key=GOOGLE_API_KEY
)

# Google Sheets Auth using Streamlit secrets for GCP service account
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_dict(dict(st.secrets["gcp"]), scope)
client = gspread.authorize(creds)
sheet = client.open("CustomerSupportLog").sheet1

# Ensure headers exist
expected_headers = ["Timestamp", "Query", "Department", "Sentiment", "Priority", "Assigned Employee", "Auto Response"]
existing_headers = sheet.row_values(1)
if existing_headers != expected_headers:
    st.error("❌ Google Sheet is not correctly set up. Please fix the headers manually and restart the app.")
    st.stop()

# Define departments and employees
employee_groups = {
    "Billing": {"High": ["Tanish"], "Medium": ["Mayank", "Yash"], "Low": ["Leena", "Ketki"]},
    "Technical Support": {"High": ["Himanshi"], "Medium": ["Kashish", "Siddhi"], "Low": ["Saloni", "Sanika"]},
    "Sales": {"High": ["Rushikesh"], "Medium": ["Om", "Nishant"], "Low": ["Sahil", "Bharat"]}
}

employee_emails = {
    "Tanish": "ay19tani@gmail.com",
    "Himanshi":"ay19tani@gmail.com",
    "Kashish":"ay19tani@gmail.com",
    "Siddhi":"ay19tani@gmail.com",
    "Saloni":"ay19tani@gmail.com",
    "Sanika":"ay19tani@gmail.com",
    "Rushikesh":"ay19tani@gmail.com",
    "Om":"ay19tani@gmail.com",
    "Nishant":"ay19tani@gmail.com",
    "Sahil":"ay19tani@gmail.com",
    "Bharat":"ay19tani@gmail.com",
    "Mayank": "mayank.mhatre22@it.sce.edu.in",
    "Yash": "yash.salunkhe22@it.sce.edu.in",
    "Leena": "leenakatkar60@gmail.com",
    "Ketki": "ketkimane1806@gmail.com"
}

# Initialize workload tracker
workload_tracker = {emp: 0 for emp in employee_emails}
records = sheet.get_all_records()
for row in records:
    assigned_emp = row.get("Assigned Employee", "")
    if assigned_emp in workload_tracker:
        workload_tracker[assigned_emp] += 1

# Core logic functions
def safe_llm_invoke(prompt):
    try:
        return llm.invoke([HumanMessage(content=prompt)]).content.strip()
    except Exception as e:
        st.error("❌ AI request failed. Reason: Possibly rate limit reached or key expired. Please try again later.")
        st.stop()

def classify_department(user_query):
    prompt = f"""You are an AI assistant for customer support. Classify this query:
"{user_query}" into Billing, Technical Support, or Sales. Just give the department name."""
    return safe_llm_invoke(prompt)

def detect_sentiment_and_priority(user_query):
    prompt = f"""Classify the sentiment (Positive, Neutral, Negative) and map to priority:
"{user_query}". Reply in format: Sentiment, Priority"""
    result = safe_llm_invoke(prompt)
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
    return safe_llm_invoke(prompt)

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

# Process form submission
if st.button("🚀 Submit"):
    if not user_email or not user_query:
        st.warning("Please fill both Email and Query!")
    else:
        dept = classify_department(user_query)
        sentiment, priority = detect_sentiment_and_priority(user_query)
        assigned = assign_employee(dept, priority)
        auto_reply = generate_auto_response(user_query)

        log_to_sheet(user_query, dept, sentiment, priority, assigned, auto_reply)

        if assigned in employee_emails:
            send_email(
                employee_emails[assigned],
                f"New Support Ticket Assigned ({priority})",
                f"You've been assigned a ticket:\n\nQuery: {user_query}\nPriority: {priority}\nDepartment: {dept}"
            )
        
        send_email(
            user_email,
            "Query Received – AI Support",
            f"Hi,\n\nWe received your query:\n\n\"{user_query}\"\n\nOur support team from {dept} will get back shortly.\n\n!"
        )

        st.success("✅ We have received your query, please check your Email")
        # st.write(f"**Department:** {dept}")
        # st.write(f"**Sentiment:** {sentiment}")
        # st.write(f"**Priority:** {priority}")
        # st.write(f"**Assigned Employee:** {assigned}")
        # st.info(f"**Auto Response:** {auto_reply}")
