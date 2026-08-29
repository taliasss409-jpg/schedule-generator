import io
import streamlit as st

from schedule_engine import build_schedule_workbook


# --------------------------------------------------
# הגדרות הדף
# --------------------------------------------------
st.set_page_config(
    page_title="סידור עבודה",
    page_icon="🌿",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# --------------------------------------------------
# CSS לעיצוב
# --------------------------------------------------
st.markdown("""
<style>

    /* הסתרת אלמנטים מיותרים */
    #MainMenu {
        visibility: hidden;
    }

    footer {
        visibility: hidden;
    }

    header {
        visibility: hidden;
    }

    /* רקע */
    .stApp {
        background-color: #F6F8F6;
    }

    /* תוכן מרכזי */
    .block-container {
        max-width: 850px;
        padding-top: 5rem;
        padding-bottom: 4rem;
    }

    /* כותרת */
    .main-title {
        text-align: center;
        color: #12372A;
        font-size: 3rem;
        font-weight: 700;
        margin-bottom: 0.4rem;
    }

    /* כותרת משנה */
    .subtitle {
        text-align: center;
        color: #61716A;
        font-size: 1.1rem;
        margin-bottom: 3rem;
    }

    /* כרטיס */
    .upload-card {
        background: white;
        border-radius: 20px;
        padding: 2.5rem;
        box-shadow: 0px 8px 30px rgba(18, 55, 42, 0.10);
        border: 1px solid #E6ECE8;
        margin-bottom: 1.5rem;
    }

    /* כותרת בתוך הכרטיס */
    .card-title {
        color: #12372A;
        font-size: 1.5rem;
        font-weight: 650;
        text-align: center;
        margin-bottom: 0.5rem;
    }

    .card-text {
        color: #6B756F;
        text-align: center;
        margin-bottom: 1.5rem;
    }

    /* כפתורים */
    .stButton > button,
    .stDownloadButton > button {
        width: 100%;
        border-radius: 10px;
        padding: 0.7rem;
        font-size: 1rem;
        font-weight: 600;
    }

    .stButton > button {
        background-color: #12372A;
        color: white;
        border: none;
    }

    .stButton > button:hover {
        background-color: #1D513D;
        color: white;
        border: none;
    }

    /* מידע על הקובץ */
    .file-success {
        background-color: #EDF6F0;
        color: #24513D;
        padding: 1rem;
        border-radius: 10px;
        text-align: center;
        margin-top: 1rem;
        margin-bottom: 1rem;
    }

    /* הודעת תחתית */
    .footer-text {
        text-align: center;
        color: #9AA39E;
        font-size: 0.85rem;
        margin-top: 3rem;
    }

</style>
""", unsafe_allow_html=True)


# --------------------------------------------------
# כותרת
# --------------------------------------------------
st.markdown(
    '<div class="main-title">🌿 סידור עבודה</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="subtitle">'
    'יצירת סידור עבודה במהירות ובפשטות'
    '</div>',
    unsafe_allow_html=True
)


# --------------------------------------------------
# כרטיס העלאת קובץ
# --------------------------------------------------
st.markdown(
    '<div class="upload-card">',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="card-title">העלאת קובץ בקשות</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="card-text">'
    'בחרי את קובץ ה־Excel המכיל את זמינות העובדים'
    '</div>',
    unsafe_allow_html=True
)


uploaded_file = st.file_uploader(
    "בחרי קובץ Excel",
    type=["xlsx"],
    label_visibility="collapsed"
)


if uploaded_file is not None:

    st.markdown(
        f'<div class="file-success">'
        f'✓ נבחר קובץ: <b>{uploaded_file.name}</b>'
        f'</div>',
        unsafe_allow_html=True
    )

    if st.button("⚙️ צור סידור עבודה", type="primary"):

        try:

            with st.spinner("מייצרת את סידור העבודה..."):

                requests_bytes = uploaded_file.getvalue()

                workbook, employees, day_cols = (
                    build_schedule_workbook(requests_bytes)
                )

                output = io.BytesIO()

                workbook.save(output)

                output.seek(0)

            st.success(
                f"🎉 סידור העבודה מוכן! "
                f"זוהו {len(employees)} עובדים."
            )

            st.download_button(
                label="⬇️ הורידי את סידור העבודה",
                data=output.getvalue(),
                file_name="סידור_עבודה.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                type="primary"
            )

        except Exception as e:

            st.error("אירעה שגיאה בעיבוד הקובץ.")

            with st.expander("פרטי השגיאה"):
                st.exception(e)


st.markdown(
    '<div class="upload-card">',
    unsafe_allow_html=True
)

st.markdown(
    """
    <div class="card-title">איך זה עובד?</div>

    <div class="card-text">
        ① העלאת קובץ הבקשות<br><br>
        ② יצירת סידור העבודה<br><br>
        ③ הורדת קובץ ה־Excel המוכן
    </div>
    """,
    unsafe_allow_html=True
)


# --------------------------------------------------
# Footer
# --------------------------------------------------
st.markdown(
    '<div class="footer-text">'
    'מערכת חכמה ליצירת סידורי עבודה'
    '</div>',
    unsafe_allow_html=True
)
