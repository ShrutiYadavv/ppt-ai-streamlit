import streamlit as st
from pptx import Presentation
from sentence_transformers import SentenceTransformer, util
from torch import device
import os

st.set_page_config(page_title="PPT Q&A", layout="centered")
st.title("📊 AI Q&A from PowerPoint Slides")

# Load model and force CPU usage
model = SentenceTransformer("all-mpnet-base-v2")
model.to(device("cpu"))  # Force CPU for Streamlit Cloud

def extract_slide_knowledge(pptx_path):
    prs = Presentation(pptx_path)
    slide_knowledge = []
    for slide in prs.slides:
        title = slide.shapes.title.text if slide.shapes.title else ""
        content = ""
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text != title:
                content += shape.text.strip() + "\n"
        if title or content:
            slide_knowledge.append(f"{title.strip()}\n{content.strip()}")
    return slide_knowledge

# File upload and question input
uploaded_file = st.file_uploader("📤 Upload a PowerPoint (.pptx) file", type=["pptx"])
question = st.text_input("❓ Enter your question")

if st.button("Get Answer"):
    if not uploaded_file or not question:
        st.warning("Please upload a file and enter a question.")
    else:
        try:
            # Save file locally
            save_path = "uploaded_ppt.pptx"
            with open(save_path, "wb") as f:
                f.write(uploaded_file.read())

            # Extract text from slides
            slides = extract_slide_knowledge(save_path)
            if not slides:
                st.error("No content found in slides.")
            else:
                slide_embeddings = model.encode(slides, convert_to_tensor=True)
                question_embedding = model.encode(question, convert_to_tensor=True)
                scores = util.cos_sim(question_embedding, slide_embeddings)
                best_idx = scores.argmax().item()
                confidence = scores[0][best_idx].item()
                best_slide = slides[best_idx]

                # Show result
                st.success("✅ Best matching answer found:")
                st.markdown(f"""**Slide Content:**  
{best_slide}""")
                st.markdown(f"**Confidence Score:** {confidence:.2f}")

        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
