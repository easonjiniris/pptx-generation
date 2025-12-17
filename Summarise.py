import pypandoc
import json
import os
from openai import OpenAI
import re


def summarise_document(docx_path: str, model_name: str, api_key: str, additional_prompt: str) -> str:
    print("\tLoading Template...")
    dir = os.path.dirname(os.path.abspath(__file__))
    slide_context_path = os.path.join(dir, "context", "slide_context.json")
    with open(slide_context_path, "r", encoding="utf-8") as f:
        slides_context = json.load(f)

    context = f"""
    You are a presentation designer.

    You are given:
    1. A document in mostly markdown format (sometimes with pseudo-tables).
    2. A list of slide templates you can use to build a presentation.
        Each template has:
        - id: identification for the slides.
        - category: for supplemental information, ignore all templates under the 'various' category.
        - description: when it should be used.
        - slots: the content areas available on that slide, with recommended limits. Ignore chart and image elements.

    Available slide templates:
    {json.dumps(slides_context, indent=2)}

    Your job:
    - Read the document.
    - Break it into a logical sequence of slides.
    - For each slide, choose the most appropriate template.
    - Fill the slots of that template with content derived from the document.
    - If a section has more content than fits one slide, create multiple slides.

    Important rules:
    - Use only the template ids that are provided in the templates list above.
    - Respect `max_items` and `max_words`.
    - For bullet points, put a newline character between bullets, you do not need to actually produce the bullet point symbol.
    - Always return valid JSON in the exact schema below and nothing else.
    - You are allowed to change the titles of the slides.
    - There may be multiple variations of the same slide but with different number of fields. Choose the most fitting one to not waste extra space.
    - Use only standard ASCII characters. Do not use special Unicode characters like arrows (→), bullets (•), em dashes (—), smart quotes (" "), or any other Unicode symbols. Use simple alternatives like "->" for arrows, "-" for bullets and dashes, and regular quotes (").
    - {additional_prompt}

    JSON schema for your response:
    {{
        "presentation_title": "<short overall title>",
        "slides": [
        {{
            "slide_index": <the exact same "id" as the chosen template>,
            "slide_title": "<the title of this slide>"
            "slots": {{
            // keys depend on the template's slot definitions, values are the string content
            }}
            "source_section": "<optional short reference to the source part of the doc>",
        }}
        ]
    }}
    """

    print(f"\tTemplate Loaded. Prompting {model_name}...")
    client = OpenAI(api_key=api_key)
    response = client.chat.completions.create(
        model=model_name,
        messages=[
            {"role": "system", "content": context},
            {"role": "user", "content": pypandoc.convert_file(
                docx_path, "md")},
        ]
    )

    response_path = os.path.join(dir, "output", "response.txt")
    with open(response_path, "w", encoding="utf-8") as f:
        f.write(response.choices[0].message.content)
    print(f"\tResponse Saved to {response_path}")

    cleaned_response = re.sub(
        r',(\s*[}\]])', r'\1', response.choices[0].message.content)
    return cleaned_response
