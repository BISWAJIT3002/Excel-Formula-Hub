import streamlit as st
import pandas as pd
import math
import random

# ---------------- PAGE CONFIG ----------------
st.set_page_config(
    page_title="Excel Formula Hub",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ---------------- CUSTOM CSS ----------------
st.markdown("""
<style>
.stApp {
    background-color: #f5f7fb;
}
.main-title {
    font-size: 38px;
    font-weight: 800;
    color: #14532d;
}
.subtitle {
    color: #64748b;
    font-size: 16px;
}
.card {
    background: white;
    padding: 22px;
    border-radius: 14px;
    border: 1px solid #e5e7eb;
    margin-bottom: 12px;
}
.card h3 {
    color: #166534;
}
.formula-name {
    color: #15803d;
    font-size: 21px;
    font-weight: bold;
}
.small-text {
    color: #64748b;
    font-size: 14px;
}
div.stButton > button {
    border-radius: 8px;
}
</style>
""", unsafe_allow_html=True)


# ---------------- FORMULA DATABASE ----------------
formulas = [
    # Math
    {"name": "SUM", "category": "Math & Trigonometry", "syntax": "=SUM(A1:A10)", "description": "Adds all numbers in a selected range.", "example": "=SUM(10,20,30)", "result": "60"},
    {"name": "PRODUCT", "category": "Math & Trigonometry", "syntax": "=PRODUCT(A1:A5)", "description": "Multiplies all numbers.", "example": "=PRODUCT(2,3,4)", "result": "24"},
    {"name": "ABS", "category": "Math & Trigonometry", "syntax": "=ABS(number)", "description": "Returns the absolute value of a number.", "example": "=ABS(-25)", "result": "25"},
    {"name": "ROUND", "category": "Math & Trigonometry", "syntax": "=ROUND(number, num_digits)", "description": "Rounds a number to a specified number of digits.", "example": "=ROUND(12.567,2)", "result": "12.57"},
    {"name": "ROUNDUP", "category": "Math & Trigonometry", "syntax": "=ROUNDUP(number, num_digits)", "description": "Rounds a number up, away from zero.", "example": "=ROUNDUP(12.341,2)", "result": "12.35"},
    {"name": "ROUNDDOWN", "category": "Math & Trigonometry", "syntax": "=ROUNDDOWN(number, num_digits)", "description": "Rounds a number down, toward zero.", "example": "=ROUNDDOWN(12.349,2)", "result": "12.34"},
    {"name": "MOD", "category": "Math & Trigonometry", "syntax": "=MOD(number, divisor)", "description": "Returns the remainder after division.", "example": "=MOD(10,3)", "result": "1"},
    {"name": "POWER", "category": "Math & Trigonometry", "syntax": "=POWER(number, power)", "description": "Raises a number to a power.", "example": "=POWER(2,3)", "result": "8"},
    {"name": "SQRT", "category": "Math & Trigonometry", "syntax": "=SQRT(number)", "description": "Returns the positive square root.", "example": "=SQRT(64)", "result": "8"},
    {"name": "INT", "category": "Math & Trigonometry", "syntax": "=INT(number)", "description": "Rounds a number down to the nearest integer.", "example": "=INT(8.9)", "result": "8"},

    # Statistical
    {"name": "AVERAGE", "category": "Statistical", "syntax": "=AVERAGE(A1:A10)", "description": "Calculates the arithmetic mean.", "example": "=AVERAGE(10,20,30)", "result": "20"},
    {"name": "MAX", "category": "Statistical", "syntax": "=MAX(A1:A10)", "description": "Returns the largest value.", "example": "=MAX(10,50,30)", "result": "50"},
    {"name": "MIN", "category": "Statistical", "syntax": "=MIN(A1:A10)", "description": "Returns the smallest value.", "example": "=MIN(10,50,30)", "result": "10"},
    {"name": "COUNT", "category": "Statistical", "syntax": "=COUNT(A1:A10)", "description": "Counts cells containing numbers.", "example": "=COUNT(10,20,30)", "result": "3"},
    {"name": "COUNTA", "category": "Statistical", "syntax": "=COUNTA(A1:A10)", "description": "Counts non-empty values.", "example": "=COUNTA(10,20,'Excel')", "result": "3"},
    {"name": "MEDIAN", "category": "Statistical", "syntax": "=MEDIAN(A1:A10)", "description": "Returns the middle value.", "example": "=MEDIAN(10,20,30)", "result": "20"},
    {"name": "LARGE", "category": "Statistical", "syntax": "=LARGE(A1:A10,1)", "description": "Returns the nth largest value.", "example": "=LARGE(10,50,30)", "result": "50"},
    {"name": "SMALL", "category": "Statistical", "syntax": "=SMALL(A1:A10,1)", "description": "Returns the nth smallest value.", "example": "=SMALL(10,50,30)", "result": "10"},

    # Logical
    {"name": "IF", "category": "Logical", "syntax": '=IF(condition, value_if_true, value_if_false)', "description": "Returns one value if a condition is true and another if false.", "example": '=IF(80>=40,"Pass","Fail")', "result": "Pass"},
    {"name": "AND", "category": "Logical", "syntax": "=AND(condition1, condition2)", "description": "Returns TRUE if all conditions are true.", "example": "=AND(10>5,20>10)", "result": "TRUE"},
    {"name": "OR", "category": "Logical", "syntax": "=OR(condition1, condition2)", "description": "Returns TRUE if at least one condition is true.", "example": "=OR(10<5,20>10)", "result": "TRUE"},
    {"name": "NOT", "category": "Logical", "syntax": "=NOT(logical)", "description": "Reverses TRUE and FALSE.", "example": "=NOT(TRUE)", "result": "FALSE"},
    {"name": "IFERROR", "category": "Logical", "syntax": "=IFERROR(value, value_if_error)", "description": "Returns an alternative value if a formula produces an error.", "example": '=IFERROR(10/0,"Error")', "result": "Error"},
    {"name": "IFS", "category": "Logical", "syntax": "=IFS(condition1, value1, condition2, value2)", "description": "Checks multiple conditions and returns the first matching result.", "example": '=IFS(90>=80,"A",90>=60,"B")', "result": "A"},

    # Text
    {"name": "LEFT", "category": "Text", "syntax": "=LEFT(text, num_chars)", "description": "Extracts characters from the left side of text.", "example": '=LEFT("EXCEL",2)', "result": "EX"},
    {"name": "RIGHT", "category": "Text", "syntax": "=RIGHT(text, num_chars)", "description": "Extracts characters from the right side of text.", "example": '=RIGHT("EXCEL",2)', "result": "EL"},
    {"name": "MID", "category": "Text", "syntax": "=MID(text, start_num, num_chars)", "description": "Extracts characters from the middle of text.", "example": '=MID("EXCEL",2,3)', "result": "XCE"},
    {"name": "LEN", "category": "Text", "syntax": "=LEN(text)", "description": "Counts the number of characters.", "example": '=LEN("EXCEL")', "result": "5"},
    {"name": "TRIM", "category": "Text", "syntax": "=TRIM(text)", "description": "Removes extra spaces from text.", "example": '=TRIM(" Excel Hub ")', "result": "Excel Hub"},
    {"name": "UPPER", "category": "Text", "syntax": "=UPPER(text)", "description": "Converts text to uppercase.", "example": '=UPPER("excel")', "result": "EXCEL"},
    {"name": "LOWER", "category": "Text", "syntax": "=LOWER(text)", "description": "Converts text to lowercase.", "example": '=LOWER("EXCEL")', "result": "excel"},
    {"name": "PROPER", "category": "Text", "syntax": "=PROPER(text)", "description": "Capitalizes the first letter of each word.", "example": '=PROPER("excel formula hub")', "result": "Excel Formula Hub"},
    {"name": "CONCAT", "category": "Text", "syntax": "=CONCAT(text1, text2)", "description": "Combines text from multiple values.", "example": '=CONCAT("Excel"," Hub")', "result": "Excel Hub"},
    {"name": "SUBSTITUTE", "category": "Text", "syntax": "=SUBSTITUTE(text, old_text, new_text)", "description": "Replaces specified text with new text.", "example": '=SUBSTITUTE("Excel Hub","Hub","Formula")', "result": "Excel Formula"},

    # Date
    {"name": "TODAY", "category": "Date & Time", "syntax": "=TODAY()", "description": "Returns the current date.", "example": "=TODAY()", "result": "Current date"},
    {"name": "NOW", "category": "Date & Time", "syntax": "=NOW()", "description": "Returns the current date and time.", "example": "=NOW()", "result": "Current date and time"},
    {"name": "DAY", "category": "Date & Time", "syntax": "=DAY(date)", "description": "Returns the day of the month.", "example": "=DAY('2026-11-06')", "result": "6"},
    {"name": "MONTH", "category": "Date & Time", "syntax": "=MONTH(date)", "description": "Returns the month number.", "example": "=MONTH('2026-11-06')", "result": "11"},
    {"name": "YEAR", "category": "Date & Time", "syntax": "=YEAR(date)", "description": "Returns the year.", "example": "=YEAR('2026-11-06')", "result": "2026"},
    {"name": "HOUR", "category": "Date & Time", "syntax": "=HOUR(time)", "description": "Returns the hour from a time value.", "example": "=HOUR('14:30')", "result": "14"},

    # Lookup
    {"name": "VLOOKUP", "category": "Lookup & Reference", "syntax": "=VLOOKUP(value, table, col_index, FALSE)", "description": "Searches for a value in the first column of a table.", "example": "=VLOOKUP(A2,A2:D10,2,FALSE)", "result": "Matching value"},
    {"name": "HLOOKUP", "category": "Lookup & Reference", "syntax": "=HLOOKUP(value, table, row_index, FALSE)", "description": "Searches for a value in the top row of a table.", "example": "=HLOOKUP(A1,A1:D5,2,FALSE)", "result": "Matching value"},
    {"name": "XLOOKUP", "category": "Lookup & Reference", "syntax": "=XLOOKUP(value, lookup_array, return_array)", "description": "Searches a range and returns a corresponding result.", "example": "=XLOOKUP(E2,A2:A10,B2:B10)", "result": "Matching value"},
    {"name": "INDEX", "category": "Lookup & Reference", "syntax": "=INDEX(array, row_num, column_num)", "description": "Returns a value at a specified position.", "example": "=INDEX(A1:C5,2,3)", "result": "Selected cell"},
    {"name": "MATCH", "category": "Lookup & Reference", "syntax": "=MATCH(value, range, 0)", "description": "Returns the relative position of a value.", "example": "=MATCH(20,A1:A5,0)", "result": "Position"},

    # Conditional
    {"name": "SUMIF", "category": "Conditional", "syntax": "=SUMIF(range, criteria, sum_range)", "description": "Adds values that meet one condition.", "example": '=SUMIF(A1:A5,"Apple",B1:B5)', "result": "Matching sum"},
    {"name": "SUMIFS", "category": "Conditional", "syntax": "=SUMIFS(sum_range, criteria_range1, criteria1)", "description": "Adds values that meet multiple conditions.", "example": "=SUMIFS(C1:C10,A1:A10,'East')", "result": "Matching sum"},
    {"name": "COUNTIF", "category": "Conditional", "syntax": "=COUNTIF(range, criteria)", "description": "Counts cells that meet one condition.", "example": '=COUNTIF(A1:A10,"Pass")', "result": "Matching count"},
    {"name": "COUNTIFS", "category": "Conditional", "syntax": "=COUNTIFS(range1,criteria1,range2,criteria2)", "description": "Counts cells that meet multiple conditions.", "example": '=COUNTIFS(A1:A10,"Pass",B1:B10,"Yes")', "result": "Matching count"},
    {"name": "AVERAGEIF", "category": "Conditional", "syntax": "=AVERAGEIF(range, criteria, average_range)", "description": "Calculates the average for values meeting a condition.", "example": '=AVERAGEIF(A1:A10,"Pass",B1:B10)', "result": "Matching average"},

    # Financial
    {"name": "PMT", "category": "Financial", "syntax": "=PMT(rate, nper, pv)", "description": "Calculates loan payment per period.", "example": "=PMT(10%/12,12,10000)", "result": "Periodic payment"},
    {"name": "FV", "category": "Financial", "syntax": "=FV(rate, nper, pmt)", "description": "Calculates the future value of an investment.", "example": "=FV(5%/12,12,-1000)", "result": "Future value"},
    {"name": "PV", "category": "Financial", "syntax": "=PV(rate, nper, pmt)", "description": "Calculates the present value of an investment.", "example": "=PV(5%/12,12,-1000)", "result": "Present value"}
]

categories = sorted(set(item["category"] for item in formulas))


# ---------------- SESSION STATE ----------------
if "quiz_submitted" not in st.session_state:
    st.session_state.quiz_submitted = False

if "quiz_questions" not in st.session_state:
    st.session_state.quiz_questions = []

if "quiz_answers" not in st.session_state:
    st.session_state.quiz_answers = {}


# ---------------- SIDEBAR ----------------
st.sidebar.markdown("## 📊 Excel Formula Hub")
st.sidebar.caption("Your Excel Learning Platform")
st.sidebar.divider()

page = st.sidebar.radio(
    "MAIN MENU",
    [
        "🏠 Home",
        "📚 All Formulas",
        "🔍 Search Formula",
        "📂 Categories",
        "🎓 Learning Courses",
        "📝 Practice Quiz",
        "🧮 Formula Playground",
        "📥 Download Resources",
        "ℹ️ About Us"
    ]
)

st.sidebar.divider()
st.sidebar.success("Learn Excel. Practice Excel. Master Excel.")


# =====================================================
# HOME PAGE
# =====================================================
if page == "🏠 Home":

    st.markdown('<p class="main-title">Welcome to Excel Formula Hub 📊</p>', unsafe_allow_html=True)
    st.markdown(
        '<p class="subtitle">Your complete platform to learn, explore and practice Microsoft Excel formulas.</p>',
        unsafe_allow_html=True
    )

    st.divider()

    c1, c2, c3 = st.columns(3)

    with c1:
        st.metric("Total Formulas", len(formulas))

    with c2:
        st.metric("Formula Categories", len(categories))

    with c3:
        st.metric("Learning Levels", 3)

    st.markdown("## 🚀 Explore Excel Formula Hub")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("""
        <div class="card">
        <h3>📚 All Formulas</h3>
        <p>Explore the formula database with syntax and explanations.</p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("""
        <div class="card">
        <h3>🎓 Learning Courses</h3>
        <p>Learn Excel from Beginner to Advanced level.</p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("""
        <div class="card">
        <h3>🧮 Formula Playground</h3>
        <p>Enter supported formulas and test their results.</p>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown("""
        <div class="card">
        <h3>🔍 Search Formula</h3>
        <p>Find formulas quickly using a name or keyword.</p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("""
        <div class="card">
        <h3>📝 Practice Quiz</h3>
        <p>Test your knowledge with Excel MCQs.</p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("""
        <div class="card">
        <h3>📥 Download Resources</h3>
        <p>Download formula lists for offline learning.</p>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("## ⭐ Popular Excel Formulas")

    popular = ["SUM", "AVERAGE", "IF", "VLOOKUP", "COUNTIF", "XLOOKUP"]

    cols = st.columns(3)

    for i, name in enumerate(popular):
        item = next(f for f in formulas if f["name"] == name)

        with cols[i % 3]:
            st.markdown(f"""
            <div class="card">
                <h3>{item['name']}</h3>
                <p>{item['description']}</p>
                <code>{item['syntax']}</code>
            </div>
            """, unsafe_allow_html=True)

    st.markdown("## 📂 Formula Categories")

    cols = st.columns(4)
    for i, category in enumerate(categories):
        with cols[i % 4]:
            count = len([f for f in formulas if f["category"] == category])
            st.info(f"**{category}**\n\n{count} formulas")

    st.divider()
    st.caption("© 2026 Excel Formula Hub | Learn Excel Easily")


# =====================================================
# ALL FORMULAS
# =====================================================
elif page == "📚 All Formulas":

    st.title("📚 All Excel Formulas")
    st.write("Browse the formula database with syntax, examples and explanations.")

    selected_category = st.selectbox(
        "Filter by Category",
        ["All Categories"] + categories
    )

    filtered = formulas

    if selected_category != "All Categories":
        filtered = [f for f in formulas if f["category"] == selected_category]

    st.write(f"**Total results: {len(filtered)}**")

    for item in filtered:
        with st.expander(f"{item['name']} — {item['category']}"):
            st.markdown(f"**Description:** {item['description']}")
            st.code(item["syntax"], language="excel")
            st.markdown(f"**Example:** `{item['example']}`")
            st.markdown(f"**Example result:** `{item['result']}`")

    df = pd.DataFrame(filtered)

    st.download_button(
        "⬇️ Download Formula List (CSV)",
        data=df.to_csv(index=False).encode("utf-8"),
        file_name="excel_formulas.csv",
        mime="text/csv"
    )


# =====================================================
# SEARCH FORMULA
# =====================================================
elif page == "🔍 Search Formula":

    st.title("🔍 Search Excel Formula")
    st.write("Search by formula name, category or description.")

    search = st.text_input(
        "Enter formula name or keyword",
        placeholder="Example: SUM, IF, text, date..."
    )

    if search:
        results = [
            f for f in formulas
            if search.lower() in f["name"].lower()
            or search.lower() in f["category"].lower()
            or search.lower() in f["description"].lower()
        ]

        if results:
            st.success(f"{len(results)} result(s) found.")

            for item in results:
                with st.container(border=True):
                    st.subheader(item["name"])
                    st.caption(item["category"])
                    st.write(item["description"])
                    st.code(item["syntax"], language="excel")
                    st.write("Example:", item["example"])
        else:
            st.warning("No matching formula found. Try another keyword.")

    else:
        st.info("Enter a formula name or keyword to begin searching.")


# =====================================================
# CATEGORIES
# =====================================================
elif page == "📂 Categories":

    st.title("📂 Excel Formula Categories")
    st.write("Explore formulas based on their category.")

    selected = st.selectbox("Choose a category", categories)

    selected_formulas = [
        f for f in formulas if f["category"] == selected
    ]

    st.subheader(selected)
    st.write(f"Total formulas: {len(selected_formulas)}")

    for item in selected_formulas:
        with st.expander(item["name"]):
            st.write(item["description"])
            st.code(item["syntax"], language="excel")
            st.write("Example:", item["example"])
            st.write("Example result:", item["result"])


# =====================================================
# LEARNING COURSES
# =====================================================
elif page == "🎓 Learning Courses":

    st.title("🎓 Excel Learning Courses")
    st.write("Choose your learning level and start learning Excel.")

    level = st.selectbox(
        "Select Learning Level",
        ["Beginner", "Intermediate", "Advanced"]
    )

    courses = {
        "Beginner": [
            ("Introduction to Excel", "Understand workbook, worksheet, rows, columns and cells."),
            ("Basic Formulas", "Learn SUM, AVERAGE, MIN, MAX and COUNT."),
            ("Basic Formatting", "Learn cell formatting, borders, colours and number formats."),
            ("Basic Data Management", "Learn sorting, filtering and basic tables.")
        ],
        "Intermediate": [
            ("Logical Formulas", "Learn IF, AND, OR, NOT and IFERROR."),
            ("Conditional Formulas", "Learn SUMIF, SUMIFS, COUNTIF and COUNTIFS."),
            ("Text Formulas", "Learn LEFT, RIGHT, MID, LEN, TRIM and CONCAT."),
            ("Lookup Formulas", "Learn VLOOKUP, HLOOKUP, INDEX and MATCH.")
        ],
        "Advanced": [
            ("Advanced Lookup", "Learn XLOOKUP and advanced lookup combinations."),
            ("Data Analysis", "Learn Pivot Tables, charts and data summaries."),
            ("Dynamic Arrays", "Explore FILTER, SORT and UNIQUE in supported Excel versions."),
            ("Financial Formulas", "Understand PMT, FV and PV.")
        ]
    }

    st.subheader(f"{level} Course")

    for index, (title, description) in enumerate(courses[level], 1):
        with st.expander(f"Lesson {index}: {title}"):
            st.write(description)

    st.info("Course topics are learning guides. Practical lesson material can be added in future updates.")


# =====================================================
# PRACTICE QUIZ
# =====================================================
elif page == "📝 Practice Quiz":

    st.title("📝 Excel Practice Quiz")
    st.write("Test your knowledge of Excel formulas.")

    questions = [
        {
            "q": "Which formula adds numbers?",
            "options": ["AVERAGE", "SUM", "COUNT", "MAX"],
            "answer": "SUM",
            "explanation": "SUM adds numbers in selected cells or ranges."
        },
        {
            "q": "Which formula calculates the average?",
            "options": ["MIN", "MAX", "AVERAGE", "COUNT"],
            "answer": "AVERAGE",
            "explanation": "AVERAGE calculates the arithmetic mean."
        },
        {
            "q": "Which formula checks a condition?",
            "options": ["IF", "SUM", "LEFT", "MAX"],
            "answer": "IF",
            "explanation": "IF returns different results based on a condition."
        },
        {
            "q": "Which formula counts numeric cells?",
            "options": ["COUNTA", "COUNT", "SUM", "LEN"],
            "answer": "COUNT",
            "explanation": "COUNT counts cells containing numbers."
        },
        {
            "q": "Which formula returns the largest value?",
            "options": ["MIN", "SMALL", "MAX", "ROUND"],
            "answer": "MAX",
            "explanation": "MAX returns the largest value in a range."
        },
        {
            "q": "Which formula counts characters in text?",
            "options": ["LEN", "LEFT", "RIGHT", "TRIM"],
            "answer": "LEN",
            "explanation": "LEN returns the number of characters."
        },
        {
            "q": "Which formula extracts characters from the left?",
            "options": ["RIGHT", "MID", "LEFT", "LOWER"],
            "answer": "LEFT",
            "explanation": "LEFT extracts characters from the beginning of text."
        },
        {
            "q": "Which formula returns the current date?",
            "options": ["NOW", "TODAY", "DAY", "YEAR"],
            "answer": "TODAY",
            "explanation": "TODAY returns the current date."
        },
        {
            "q": "Which formula searches vertically in a table?",
            "options": ["HLOOKUP", "VLOOKUP", "SUMIF", "COUNT"],
            "answer": "VLOOKUP",
            "explanation": "VLOOKUP searches in the first column of a table."
        },
        {
            "q": "Which formula rounds a number?",
            "options": ["ABS", "MOD", "ROUND", "POWER"],
            "answer": "ROUND",
            "explanation": "ROUND rounds a number to a specified number of digits."
        },
        {
            "q": "Which formula returns the smallest value?",
            "options": ["MAX", "MIN", "LARGE", "SUM"],
            "answer": "MIN",
            "explanation": "MIN returns the smallest value."
        },
        {
            "q": "Which formula combines text?",
            "options": ["CONCAT", "COUNT", "MOD", "INT"],
            "answer": "CONCAT",
            "explanation": "CONCAT combines text from multiple values."
        },
        {
            "q": "Which formula removes extra spaces?",
            "options": ["UPPER", "TRIM", "LEN", "MID"],
            "answer": "TRIM",
            "explanation": "TRIM removes extra spaces from text."
        },
        {
            "q": "Which formula returns the remainder?",
            "options": ["POWER", "MOD", "ABS", "ROUND"],
            "answer": "MOD",
            "explanation": "MOD returns the remainder after division."
        },
        {
            "q": "Which formula counts cells meeting a condition?",
            "options": ["SUMIF", "COUNTIF", "AVERAGE", "MAX"],
            "answer": "COUNTIF",
            "explanation": "COUNTIF counts cells that meet a specified condition."
        }
    ]

    if not st.session_state.quiz_questions:
        st.session_state.quiz_questions = random.sample(questions, 10)

    if not st.session_state.quiz_submitted:

        with st.form("quiz_form"):
            for i, question in enumerate(st.session_state.quiz_questions):
                st.markdown(f"### Q{i+1}. {question['q']}")

                st.session_state.quiz_answers[i] = st.radio(
                    "Choose your answer:",
                    question["options"],
                    key=f"question_{i}",
                    index=None
                )

            submitted = st.form_submit_button("Submit Quiz")

        if submitted:
            st.session_state.quiz_submitted = True
            st.rerun()

    else:

        score = sum(
            1 for i, question in enumerate(st.session_state.quiz_questions)
            if st.session_state.quiz_answers.get(i) == question["answer"]
        )

        st.subheader("Your Quiz Result")

        st.metric("Your Score", f"{score}/10")

        if score >= 8:
            st.success("Excellent! You have a strong understanding of Excel formulas.")
        elif score >= 5:
            st.info("Good work! Keep practicing.")
        else:
            st.warning("Keep learning. You can improve your score.")

        st.progress(score / 10)

        st.subheader("Answer Review")

        for i, question in enumerate(st.session_state.quiz_questions):
            user_answer = st.session_state.quiz_answers.get(i)

            st.markdown(f"**Q{i+1}. {question['q']}**")
            st.write("Your answer:", user_answer if user_answer else "Not answered")
            st.write("Correct answer:", question["answer"])

            if user_answer == question["answer"]:
                st.success("Correct")
            else:
                st.error("Incorrect")

            st.caption(question["explanation"])
            st.divider()

        if st.button("🔄 Restart Quiz"):
            st.session_state.quiz_submitted = False
            st.session_state.quiz_questions = random.sample(questions, 10)
            st.session_state.quiz_answers = {}
            st.rerun()


# =====================================================
# FORMULA PLAYGROUND
# =====================================================
elif page == "🧮 Formula Playground":

    st.title("🧮 Excel Formula Playground")
    st.write("Enter a supported formula and calculate its result.")

    st.info(
        "This is a learning simulator for selected Excel functions, "
        "not a complete Microsoft Excel calculation engine."
    )

    formula_input = st.text_input(
        "Enter Formula",
        placeholder="Example: =SUM(10,20,30)"
    )

    def calculate_formula(expression):

        expression = expression.strip()

        if expression.startswith("="):
            expression = expression[1:]

        # Only allow explicitly supported functions
        allowed = {
            "SUM": lambda *x: sum(x),
            "AVERAGE": lambda *x: sum(x) / len(x) if x else 0,
            "MAX": lambda *x: max(x),
            "MIN": lambda *x: min(x),
            "COUNT": lambda *x: len(x),
            "PRODUCT": lambda *x: math.prod(x),
            "ABS": lambda x: abs(x),
            "ROUND": lambda x, n=0: round(x, int(n)),
            "POWER": lambda x, y: x ** y,
            "SQRT": lambda x: math.sqrt(x),
            "MOD": lambda x, y: x % y
        }

        import re

        match = re.fullmatch(
            r"\s*([A-Za-z]+)\s*\((.*)\)\s*",
            expression
        )

        if not match:
            raise ValueError("Invalid formula format.")

        function_name = match.group(1).upper()
        arguments_text = match.group(2)

        if function_name not in allowed:
            raise ValueError("This function is not supported in the playground.")

        if not arguments_text.strip():
            args = []
        else:
            parts = arguments_text.split(",")
            args = []

            for part in parts:
                part = part.strip()

                try:
                    args.append(float(part))
                except ValueError:
                    raise ValueError(
                        "Please enter numeric arguments only, separated by commas."
                    )

        result = allowed[function_name](*args)

        if isinstance(result, float) and result.is_integer():
            result = int(result)

        return result

    if st.button("Calculate Result", type="primary"):
        if not formula_input.strip():
            st.warning("Please enter a formula.")
        else:
            try:
                result = calculate_formula(formula_input)
                st.success(f"Result: {result}")
            except Exception as error:
                st.error(f"Unable to calculate: {error}")

    st.subheader("Supported Functions")

    st.code(
        "=SUM(10,20,30)\n"
        "=AVERAGE(10,20,30)\n"
        "=MAX(10,20,30)\n"
        "=MIN(10,20,30)\n"
        "=COUNT(10,20,30)\n"
        "=PRODUCT(2,3,4)\n"
        "=ABS(-25)\n"
        "=ROUND(12.567,2)\n"
        "=POWER(2,3)\n"
        "=SQRT(64)\n"
        "=MOD(10,3)",
        language="excel"
    )


# =====================================================
# DOWNLOAD RESOURCES
# =====================================================
elif page == "📥 Download Resources":

    st.title("📥 Download Excel Resources")
    st.write("Download the formula database for offline learning.")

    df = pd.DataFrame(formulas)

    st.subheader("Complete Formula Database")

    st.download_button(
        "⬇️ Download All Formulas",
        data=df.to_csv(index=False).encode("utf-8"),
        file_name="excel_formula_database.csv",
        mime="text/csv",
        type="primary"
    )

    st.divider()

    st.subheader("Download by Category")

    selected_category = st.selectbox(
        "Select category to download",
        categories
    )

    category_data = [
        f for f in formulas if f["category"] == selected_category
    ]

    category_df = pd.DataFrame(category_data)

    st.write(f"{len(category_data)} formulas available.")

    st.download_button(
        "⬇️ Download Selected Category",
        data=category_df.to_csv(index=False).encode("utf-8"),
        file_name=f"{selected_category.replace(' ', '_')}.csv",
        mime="text/csv"
    )

    st.divider()

    st.subheader("Formula Database Preview")
    st.dataframe(
        df[["name", "category", "syntax"]],
        use_container_width=True,
        hide_index=True
    )


# =====================================================
# ABOUT US
# =====================================================
elif page == "ℹ️ About Us":

    st.title("ℹ️ About Excel Formula Hub")

    st.markdown("""
    ### 📊 What is Excel Formula Hub?

    Excel Formula Hub is a learning platform designed to help students,
    beginners and professionals understand Microsoft Excel formulas.

    ### 🎯 Our Mission

    - Make Excel learning simple.
    - Provide formula syntax and explanations.
    - Help users practice Excel formulas.
    - Support students and job seekers in developing spreadsheet skills.

    ### 🌟 Main Features

    - Formula database
    - Formula search
    - Category-wise learning
    - Learning courses
    - Practice quizzes
    - Formula playground
    - Downloadable resources

    ### 👨‍💻 Technology Used

    - Python
    - Streamlit
    - Pandas

    ### 📌 Disclaimer

    This website is an independent educational project.
    It is not affiliated with or endorsed by Microsoft.
    The formula database is a starter collection of commonly used Excel formulas,
    not an exhaustive list of every Excel function.

    """)

    st.divider()
    st.markdown("**Excel Formula Hub | Learn Excel Easily**")
    st.caption("© 2026 Excel Formula Hub. Educational project.")
