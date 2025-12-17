import os
import json
from SlideBuilder import SlideBuilder
from Summarise import summarise_document
from dotenv import load_dotenv


def _get_response(docx_path: str, model_name: str, api_key: str, additional_prompt: str):
    print("Summarising Document...")
    response = summarise_document(
        docx_path, model_name, api_key, additional_prompt)

    return json.loads(response)


def _build_slide(data: str) -> None:
    dir = os.path.dirname(os.path.abspath(__file__))
    library_path = os.path.join(
        dir, "context", "Slides_library Argon & Co.pptx")
    output_path = os.path.join(dir, "output", "output.pptx")

    print("Summary Complete. Extracting Data...")
    presentation_title = data['presentation_title']
    slides = data['slides']

    builder = SlideBuilder(library_path, output_path)

    print(f"Creating {presentation_title}...")
    builder.create_blank()

    slides_to_copy = [0] + [slide["slide_index"]-1 for slide in slides]

    print("Copying Slides " + ", ".join(map(str, slides_to_copy)) + "...")
    builder.copy_slides(slides_to_copy)
    builder.fill_slide_type_title(presentation_title)
    builder.save_output()

    for i in range(len(slides)):
        slide = slides[i]
        slide_index = slide["slide_index"]
        print(f"Editing Slide {i+2} (Type {slide_index})...")
        method_name = f"fill_slide_type_{slide_index}"
        if not hasattr(builder, method_name):
            raise ValueError(f"Unknown Slide Type: {slide_index}")
        method = getattr(builder, method_name)
        # +2 to be 1-based and add title slide
        method(i+2, slide["slide_title"], slide["slots"])
        builder.save_output()

    print("Presentation Edited and Saved.")
    builder.close_output()
    print("Closing PowerPoint...")
    SlideBuilder.quit_powerpoint()
    print(f"PowerPoint Closed. Output Saved to {output_path}")


def generate_slide(docx_path: str) -> None:
    load_dotenv(override=True)
    model_name = os.getenv("MODEL_NAME")
    api_key = os.getenv("OPENAI_API_KEY")
    additional_prompt = os.getenv("ADDITIONAL_PROMPT")
    if not model_name:
        raise Exception("Please select a model!")
    if not api_key:
        raise Exception("Please enter an API key!")
    _build_slide(_get_response(
        docx_path, model_name, api_key, additional_prompt))


def regenerate_slide(response_path: str = "response.txt") -> None:
    with open(response_path, 'r', encoding='utf-8') as f:
        _build_slide(json.load(f))
