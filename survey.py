import streamlit as st
import pandas as pd
import os

# Ensure openpyxl is installed for Excel operations
import subprocess
import sys

try:
    import openpyxl
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
    import openpyxl

# Load or create Excel file
def load_excel(filepath):
    if os.path.exists(filepath):
        df = pd.read_excel(filepath)
    else:
        # Create a new DataFrame with expected columns
        df = pd.DataFrame(columns=[
            'Major', 'Academic Standing', 'Applied Before', 'Difficulty (1-5)',
            'Biggest Obstacle', 'Confidence in Resume', 'Attended Career Event',
            'Helpful Strategies', 'Important Skill', 'Advice'
        ])
    return df

# Save DataFrame to Excel
def save_to_excel(df, filepath):
    df.to_excel(filepath, index=False)

# Path to your Excel file
EXCEL_FILE = 'Survey_Responses.xlsx'

df = load_excel(EXCEL_FILE)

st.title("STEM Students Job Hunting Survey")
st.divider()
st.write("Please fill out the following survey.")
st.divider()

# --- Survey Questions ---
st.text_input("Please enter your name")

major = st.selectbox("What is your current major?", [
    "Computer Science", "Engineering", "Mathematics", "Physical Sciences", "Other"])

academic_standing = st.selectbox("What is your current academic standing?", [
    "Freshman", "Sophomore", "Junior", "Senior", "Graduate Student"])

applied_before = st.radio("Have you previously applied for a job or internship related to your STEM field?", ["Yes", "No"])

difficulty = st.slider("On a scale of 1 to 5, how difficult do you find it to secure a job or internship in your field?", 1, 5)

obstacles = st.multiselect("What is the biggest obstacle you have faced while searching for a job or internship?", [
    "Lack of experience", "High competition", "Lack of networking opportunities",
    "Poor interviewing skills", "Limited job openings in my field", "Other"])

confidence = st.selectbox("How confident are you in your resume and cover letter writing skills?", [
    "Very confident", "Somewhat confident", "Neutral", "Somewhat unconfident", "Very unconfident"])

attended_event = st.radio("Have you ever attended a career fair, networking event, or workshop to aid in your job search?", ["Yes", "No"])

strategies = st.multiselect("Which strategies have you found most helpful in your job or internship search?", [
    "Using LinkedIn or other job platforms", "Networking with professionals/alumni",
    "University career services", "Cold emailing companies", "Attending career fairs or workshops", "Getting referrals"])

important_skill = st.text_input("In your opinion, what skill is most important for securing a job or internship in your STEM field?")

advice = st.text_area("What advice would you give to other STEM students currently job hunting?")

# --- Submit Button ---
if st.button("Submit Survey"):
    new_response = {
        'Major': major,
        'Academic Standing': academic_standing,
        'Applied Before': applied_before,
        'Difficulty (1-5)': difficulty,
        'Biggest Obstacle': ", ".join(obstacles),
        'Confidence in Resume': confidence,
        'Attended Career Event': attended_event,
        'Helpful Strategies': ", ".join(strategies),
        'Important Skill': important_skill,
        'Advice': advice
    }

    df = pd.concat([df, pd.DataFrame([new_response])], ignore_index=True)
    save_to_excel(df, EXCEL_FILE)

    st.success("Thank you for completing the survey!")
    st.balloons()
