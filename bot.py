import streamlit as st
from docx import Document
from pptx import Presentation
import io

st.set_page_config(page_title="Hujjat Constructor", page_icon="📝")

st.title("📝 Ideal Hujjat Constructor")
st.info("Ma'lumotlarni kiriting, men ularni Word yoki PowerPoint-ga chiroyli joylab beraman.")

# Sozlamalar
with st.sidebar:
    st.header("⚙️ Sozlamalar")
    limit = st.slider("Necha varoq/slayd bo'lsin?", 1, 10, 3)
    app_type = st.selectbox("Ilova turi", ["Word", "PowerPoint"])

# Ma'lumot yig'ish
st.write(f"### {limit} ta varoq uchun ma'lumotlarni to'ldiring:")
content_data = []

for i in range(limit):
    with st.expander(f"{i+1}-chi varoq ma'lumoti", expanded=True):
        title = st.text_input(f"Sarlavha {i+1}", key=f"t{i}")
        text = st.text_area(f"Asosiy matn {i+1}", key=f"m{i}")
        image = st.file_uploader(f"Rasm (ixtiyoriy) {i+1}", type=['jpg', 'png', 'jpeg'], key=f"img{i}")
        content_data.append({"title": title, "text": text, "image": image})

# Fayl yaratish tugmasi
if st.button("🚀 Hujjatni tayyorlash"):
    if not any(d['title'] for d in content_data):
        st.warning("Iltimos, kamida bitta sarlavha kiriting!")
    else:
        if app_type == "Word":
            doc = Document()
            for d in content_data:
                doc.add_heading(d['title'], 0)
                doc.add_paragraph(d['text'])
                if d['image']:
                    doc.add_picture(d['image'], width=io.BytesIO(d['image'].read()).getbuffer().nbytes) # Soddalashtirilgan
                doc.add_page_break()
            
            output = io.BytesIO()
            doc.save(output)
            st.download_button("📥 Word faylni yuklab olish", output.getvalue(), "hujjat.docx")

        else: # PowerPoint
            prs = Presentation()
            for d in content_data:
                slide = prs.slides.add_slide(prs.slide_layouts[1])
                slide.shapes.title.text = d['title']
                slide.placeholders[1].text = d['text']
            
            output = io.BytesIO()
            prs.save(output)
            st.download_button("📥 PPTX faylni yuklab olish", output.getvalue(), "prezentatsiya.pptx")
