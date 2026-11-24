Aha! Screenshot mein saaf nazar aa raha hai ke `1. Clone the repository:` wale box ke neeche ek extra ````` aa gaya hai (aur URL ke gird `[...]` brackets lag gaye hain), is liye neeche wala part (`2. **Install required libraries:**`) raw text ban gaya hai.

Isay **10 second** mein fix karne ke liye:

1. GitHub par `README.md` ke **Edit (pencil)** icon par click karein.
2. **`Ctrl + A`** aur **`Backspace`** daba kar poora purana text saaf kar dein.
3. Neeche diya gaya **exact clean text** copy-paste karke **Commit changes** daba dein:

# Study Buddy AI

**Study Buddy AI** is an AI-powered learning and revision web application built with **Python**, **Streamlit**, and the **Google Gemini API**. Students can upload lecture notes in PDF, Word (`.docx`), PowerPoint (`.pptx`), or text (`.txt`) format to get context-grounded tutoring, non-spoiler guided hints, active-recall flashcards, practice quizzes, and spaced revision analytics.

## Key Features

* **Multi-Format File Support:** Extracts text from PDF (including scanned PDFs via Gemini Vision OCR), Word (`.docx`), PowerPoint (`.pptx`), and `.txt` files.
* **Ask & Guided Tutor:** Answers questions strictly from uploaded study material with conversation memory, plus a **Guided Hint Mode** that gives short thinking clues without spoiling the answer.
* **AI Flashcards:** Generates up to 30 flip cards with progressive clues and deck shuffling.
* **Practice Quiz:** Creates customizable multiple-choice quizzes with optional hints, instant grading, and explanations.
* **Spaced Revision & Analytics:** Tracks quiz scores and schedules spaced-repetition review dates (1, 3, and 7 days) based on performance.
* **Student Accounts & Workspace Persistence:** Secure SHA-256 login authentication with auto-saved workspaces.

## Installation & Setup

1. **Clone the repository:**
`git clone https://github.com/payal-codes-web/Study-Buddy-AI.git`
`cd Study-Buddy-AI`
2. **Install required libraries:**
`pip install -r requirements.txt`
3. **Configure the Gemini API key:**
Create a `.env` file in the root folder and add your API key:
`GEMINI_API_KEY=your_gemini_api_key_here`

## Execution (Running the App)

Run the following command in your terminal:
`streamlit run app.py`
Then open `http://localhost:8501` in your browser.

## Group Members — Salim Habib University (SHU)

* Payal
* Wania
* Farwa
