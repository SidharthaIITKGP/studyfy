import streamlit as st
from langchain_google_genai import ChatGoogleGenerativeAI
from langchain_core.messages import HumanMessage
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import io
import base64
import os

# --- 1. Caching Functions ---

@st.cache_resource
def init_llm(api_key):
    """Initializes the Gemini LLM."""
    return ChatGoogleGenerativeAI(
        model="gemini-2.5-pro",
        google_api_key=api_key
    )

@st.cache_data(show_spinner=False)
def parse_ppt_multimodal(file_bytes):
    """
    Parses a PowerPoint file from bytes, extracting text and images
    from each slide, including grouped and placeholder shapes.
    """
    pptx_file = io.BytesIO(file_bytes)
    prs = Presentation(pptx_file)
    slides_data = []

    for i, slide in enumerate(prs.slides):
        slide_text = []
        slide_images = []
        
        def find_shapes(shapes):
            for shape in shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            slide_text.append(run.text)
                if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                    find_shapes(shape.shapes)
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    try:
                        slide_images.append(shape.image.blob)
                    except Exception: pass
                if shape.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER:
                    if hasattr(shape, 'image') and shape.image:
                        try:
                            slide_images.append(shape.image.blob)
                        except Exception: pass

        find_shapes(slide.shapes)
        slides_data.append({
            "slide_number": i + 1,
            "text": "\n".join(slide_text).strip(),
            "images": slide_images
        })
        
    return slides_data

# --- 2. Core LLM Generation Function ---

def get_gemini_explanation(llm, slide_data, detail_level, include_images, prev_slide_text=None):
    """
    Generates a detailed explanation for a single slide's data
    based on the user's selected options.
    """
    
    # 1. Define the prompt based on the user's desired detail level
    detail_prompts = {
        "Summary": "Provide a concise, high-level summary of the following slide content (max 3 sentences).",
        "Standard": "Explain the key concepts on this slide, defining important terms. Keep it clear and focused.",
        "In-Depth": (
            "You are an expert university professor. Your task is to explain the content of this "
            "PowerPoint slide in great detail for a student preparing for an exam.\n"
            "You MUST combine all information (text and images, if provided) into a single, comprehensive explanation.\n"
            "1.  **Analyze the text** and **define all key terms**.\n"
            "2.  **Analyze any images** (diagrams, charts, graphs) and explain what they show.\n"
            "3.  **Connect the text to the images.** (e.g., 'The text mentions X, which is labeled in the diagram...').\n"
            "4.  **Explain the 'why'** behind the concepts. Why is this important?"
        )
    }
    
    instruction_text = detail_prompts[detail_level]

    # 2. Add previous slide context if provided
    if prev_slide_text:
        instruction_text += (
            f"\n\n--- PREVIOUS SLIDE CONTEXT ---\n"
            f"For context, here is the text from the PREVIOUS slide. Use this to inform your explanation "
            f"of the current slide, but do not explain the previous slide itself.\n"
            f"{prev_slide_text}\n"
            f"--- END PREVIOUS SLIDE CONTEXT ---"
        )
    
    # 3. Add the current slide's text
    instruction_text += f"\n\n--- CURRENT SLIDE CONTENT ---"
    if slide_data["text"]:
        instruction_text += f"\n**Slide Text:**\n{slide_data['text']}"
    else:
        instruction_text += "\n[This slide contains no text.]"

    # 4. Prepare the final message parts (text + images)
    message_parts = [{"type": "text", "text": instruction_text}]
    
    # Add images ONLY if the user has ticked the box
    if include_images and slide_data["images"]:
        message_parts.append({"type": "text", "text": "\n**Slide Images:**\n(See attached images)"})
        for img_bytes in slide_data["images"]:
            try:
                b64_image = base64.b64encode(img_bytes).decode('utf-8')
                message_parts.append({
                    "type": "image_url",
                    "image_url": {"url": f"data:image/png;base64,{b64_image}"}
                })
            except Exception as e:
                print(f"Error encoding image for API: {e}")
    
    # 5. Create and invoke the LangChain message
    try:
        human_message = HumanMessage(content=message_parts)
        response = llm.invoke([human_message])
        return response.content
    except Exception as e:
        st.error(f"An error occurred while calling the Gemini API: {str(e)}")
        return f"Error processing slide {slide_data['slide_number']}: {str(e)}"

# --- 3. Main Streamlit App ---

def main():
    st.set_page_config(
        page_title="Studyfy v5",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    st.sidebar.title("Studyfy 5.0")

    # 1. API Key (works locally with .streamlit/secrets.toml)
    api_key = st.sidebar.text_input("Enter your Google (Gemini) API Key:", type="password")

    # 2. File Uploader
    uploaded_file = st.sidebar.file_uploader(
        "Upload your PowerPoint", 
        type="pptx",
        help="Upload a .pptx file to begin."
    )
    
    # --- Session State Initialization ---
    if 'explanations' not in st.session_state:
        # 'explanations' is now a dictionary.
        # The KEY will be a unique string based on the settings (e.g., "slide_5_In-Depth_True_True")
        # The VALUE will be the generated text.
        st.session_state.explanations = {}
    if 'parsed_slides' not in st.session_state:
        st.session_state.parsed_slides = []
    if 'current_file_id' not in st.session_state:
        st.session_state.current_file_id = None

    # --- 3. Handle File Upload ---
    if not uploaded_file:
        st.info("Upload a .pptx file in the sidebar to get started.")
        return

    if not api_key:
        st.error("GOOGLE_API_KEY not found in Streamlit Secrets.")
        return

    # Check if this is a new file. If so, clear old data.
    file_id = uploaded_file.file_id
    if st.session_state.current_file_id != file_id:
        with st.spinner("Parsing presentation..."):
            file_bytes = uploaded_file.getvalue()
            st.session_state.parsed_slides = parse_ppt_multimodal(file_bytes)
            st.session_state.explanations = {} # Clear old explanations
            st.session_state.current_file_id = file_id
            st.rerun() # Rerun to update the radio button

    if not st.session_state.parsed_slides:
        st.error("There was an error parsing the presentation. Please try again.")
        return
        
    # --- 4. ADDED BACK: Analysis Options ---
    st.sidebar.header("Analysis Options")
    
    detail_level = st.sidebar.radio(
        "Level of Detail",
        ["Summary", "Standard", "In-Depth"],
        index=2, # Default to "In-Depth"
        help="Choose how detailed you want the explanation to be."
    )
    
    include_images = st.sidebar.checkbox(
        "Analyze Images (Multimodal)",
        value=True,
        help="Allow the AI to 'see' and analyze the images on the slide."
    )
    
    include_context = st.sidebar.checkbox(
        "Include Previous Slide Context",
        value=True,
        help="Provide text from the previous slide to the AI for better context."
    )

    # --- 5. Sidebar Navigation ---
    st.sidebar.header("Slide Navigation")
    slide_titles = [f"Slide {s['slide_number']}" for s in st.session_state.parsed_slides]
    
    selected_index = st.sidebar.radio(
        "Select a slide:",
        options=range(len(slide_titles)),  # Use index as the value
        format_func=lambda x: slide_titles[x], # Show "Slide X" as the label
        key='selected_index'
    )
    
    current_slide = st.session_state.parsed_slides[selected_index]
    
    # --- 6. Main Page Display ---
    
    st.title(f"Slide {current_slide['slide_number']}")
    
    # --- Raw Content Expander ---
    with st.expander("View Raw Slide Content", expanded=False):
        st.markdown("**Extracted Text:**")
        if current_slide["text"]:
            st.text(current_slide["text"])
        else:
            st.info("This slide contains no extractable text.")
        
        st.markdown("**Extracted Images:**")
        if current_slide["images"]:
            cols = st.columns(min(len(current_slide["images"]), 3)) 
            for i, img_bytes in enumerate(current_slide["images"]):
                cols[i % 3].image(
                    img_bytes, 
                    use_container_width=True,
                    caption=f"Image {i+1}"
                )
        else:
            st.info("This slide contains no extractable images.")

    # --- 7. Smart "Generate-on-Select" Logic ---
    st.markdown("---")
    st.header(f"💡 {detail_level} Explanation")
    
    # Create a unique key based on slide AND settings
    explanation_key = (
        f"{st.session_state.current_file_id}_{selected_index}_"
        f"{detail_level}_{include_images}_{include_context}"
    )
    
    # Check if we have ALREADY generated this exact explanation
    if explanation_key in st.session_state.explanations:
        # Yes: Just display it instantly
        st.markdown(st.session_state.explanations[explanation_key])
    else:
        # No: Generate it now, save it, and then display it
        with st.spinner(f"Generating {detail_level} explanation for slide {current_slide['slide_number']}..."):
            llm = init_llm(api_key)
            
            # Get previous slide text for context (if toggled)
            prev_slide_text = None
            if include_context and selected_index > 0:
                prev_slide_text = st.session_state.parsed_slides[selected_index - 1]["text"]
            
            # Call the LLM with all the options
            explanation = get_gemini_explanation(
                llm,
                current_slide,
                detail_level,
                include_images,
                prev_slide_text
            )
            
            # Save to session state for next time
            st.session_state.explanations[explanation_key] = explanation
            
            # Display it
            st.markdown(explanation)

# Run the app
if __name__ == "__main__":
    main()