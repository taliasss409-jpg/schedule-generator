import io
import streamlit as st

from schedule_engine import build_schedule_workbook


# --------------------------------------------------
# הגדרות הדף
# --------------------------------------------------
st.set_page_config(
    page_title="כלי עזר לשיבוץ",
    page_icon="🌿",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# --------------------------------------------------
# CSS לעיצוב
# --------------------------------------------------
st.markdown("""
<style>

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

    /* כותרת ראשית */
    .main-title {
        text-align: center;
        direction: rtl;
        color: #12372A;
        font-size: 3rem;
        font-weight: 700;
        margin-bottom: 3rem;
    }

    /* כרטיס העלאה */
    .upload-card {
        background: white;
        border-radius: 20px;
        padding: 2.5rem;
        box-shadow: 0px 8px 30px rgba(18, 55, 42, 0.10);
        border: 1px solid #E6ECE8;
        margin-bottom: 2rem;
    }

    /* טקסט ראשי בכרטיס */
    .upload-text {
        color: #12372A;
        font-size: 1.5rem;
        font-weight: 650;
        text-align: center;
        direction: rtl;
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
        direction: rtl;
        margin-top: 1rem;
        margin-bottom: 1rem;
    }

    /* כרטיס איך זה עובד */
    .info-card {
        background: white;
        border-radius: 20px;
        padding: 2rem;
        box-shadow: 0px 8px 30px rgba(18, 55, 42, 0.06);
        border: 1px solid #E6ECE8;
        margin-top: 1rem;
    }

    .info-title {
        color: #12372A;
        font-size: 1.5rem;
        font-weight: 650;
        text-align: center;
        direction: rtl;
        margin-bottom: 1.5rem;
    }

    .info-text {
        color: #6B756F;
        text-align: center;
        direction: rtl;
        font-size: 1rem;
        line-height: 2.2;
    }

    /* Footer */
    .footer-text {
        text-align: center;
        color: #9AA39E;
        font-size: 0.85rem;
        direction: rtl;
        margin-top: 3rem;
    }

</style>
""", unsafe_allow_html=True)


# --------------------------------------------------
# כותרת ראשית
# --------------------------------------------------
st.markdown(
    '<div class="main-title">כלי עזר לשיבוץ 🌿</div>',
    unsafe_allow_html=True
)


# --------------------------------------------------
# כרטיס העלאת הקובץ
# --------------------------------------------------
st.markdown(
    '<div class="upload-card">',
    unsafe_allow_html=True
)

st.markdown(
    """
    <div class="upload-text">
        בחר את קובץ ה־<span dir="ltr">Excel</span>
        המכיל את זמינות העובדים
    </div>
    """,
    unsafe_allow_html=True
)


uploaded_file = st.file_uploader(
    "בחר קובץ",
    type=["xlsx"],
    label_visibility="collapsed"
)


if uploaded_file is not None:

    st.markdown(
        f"""
        <div class="file-success">
            ✓ נבחר קובץ: <b>{uploaded_file.name}</b>
        </div>
        """,
        unsafe_allow_html=True
    )

    if st.button("⚙️ צור סידור זמינות", type="primary"):

        try:

            with st.spinner("מייצר את סידור העובדים..."):


                requests_bytes = uploaded_file.getvalue()

                workbook, employees, day_cols = (
                    build_schedule_workbook(requests_bytes)
                )

                output = io.BytesIO()

                workbook.save(output)

                output.seek(0)

            st.success(
                f"🎉 סידור הזמינות מוכן! זוהו {len(employees)} עובדים."
            )

            st.download_button(
                label="⬇️ הורד את קובץ זמינות העובדים",
                data=output.getvalue(),
                file_name="סידור_זמינות.xlsx",
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


# --------------------------------------------------
# איך זה עובד
# --------------------------------------------------
st.markdown(
    '<div class="info-card">',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="info-title">איך זה עובד?</div>',
    unsafe_allow_html=True
)

st.markdown(
    """
    <div class="info-text">
        ① העלאת קובץ זמינות העובדים<br>
        ② ניתוח ועיבוד הזמינות<br>
        ③ הפקת בסיס לסידור העבודה
    </div>
    """,
    unsafe_allow_html=True
)
