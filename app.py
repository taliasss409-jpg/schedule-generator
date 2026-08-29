import io
import streamlit as st

from schedule_engine import build_schedule_workbook

st.set_page_config(
    page_title="מחולל סידור עבודה",
    page_icon="📅",
    layout="wide",
)

st.title("📅 מחולל סידור עבודה")
st.write("העלי קובץ בקשות Excel והמערכת תיצור עבורך סידור עבודה ראשוני.")

uploaded_file = st.file_uploader(
    "בחרי קובץ Requests.xlsx",
    type=["xlsx"],
)

if uploaded_file is not None:
    st.success(f"נבחר קובץ: {uploaded_file.name}")

    if st.button("⚙️ צור סידור עבודה", type="primary"):
        try:
            with st.spinner("מעבד את קובץ הבקשות..."):
                requests_bytes = uploaded_file.getvalue()
                workbook, employees, day_cols = build_schedule_workbook(
                    requests_bytes
                )

                output = io.BytesIO()
                workbook.save(output)
                output.seek(0)

            st.success(
                f"✅ הסידור מוכן! זוהו {len(employees)} עובדים ו-{len(day_cols)} ימים."
            )

            st.download_button(
                label="⬇️ הורידי את סידור העבודה",
                data=output.getvalue(),
                file_name="סידור_עבודה.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )

        except Exception as e:
            st.error(f"שגיאה בעיבוד הקובץ: {e}")
            st.exception(e)
else:
    st.info("👆 התחילי בבחירת קובץ Excel.")
