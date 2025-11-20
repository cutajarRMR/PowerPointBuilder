"""
create_presentation.py
----------------------
Generates a PowerPoint deck automatically using OpenAI GPT models and your company template.

Usage:
    python create_presentation.py --topic "DataCamp Overview" --slides 8 --template "CompanyTemplate.pptx"
"""

import argparse
import json
from pathlib import Path
from openai import OpenAI
from pptx import Presentation
from pptx.util import Pt
import os
from dotenv import load_dotenv
load_dotenv()

# ---------- CONFIG ----------
MODEL = "gpt-4o-mini"  # fast + cost-effective
OUTPUT_FILE = "Generated_Presentation.pptx"
# ----------------------------

from langchain_openai import ChatOpenAI
from langchain_core.prompts import ChatPromptTemplate
import json

def generate_slide_outline(topic: str, n_slides: int, instructions: str):
    """Generate a slide outline using LangChain + OpenAI, incorporating user instructions."""

    llm = ChatOpenAI(
        model="gpt-4o-mini",
        temperature=1,
        model_kwargs={"response_format": {"type": "json_object"}} 
    )

    prompt_template = ChatPromptTemplate.from_template("""
    You are creating a professional internal PowerPoint presentation about "{topic}".
    Produce {n_slides} slides in **JSON** format.

    The user has provided the following detailed instructions for what they want in the slides:
    ---
    {instructions}
    ---
    Ensure every slide aligns with these instructions.

    Each slide must include:
    - "title": short slide title
    - "bullets": 3–5 concise bullet points
    - "notes": 2–3 sentence speaker notes explaining how to present the slide

    Output format example:
    [
      {{
        "title": "What is DataCamp?",
        "bullets": ["Online platform for data skills", "Python, R, SQL, Power BI", "Used by 10M+ learners"],
        "notes": "Introduce DataCamp as a flexible platform..."
      }}
    ]

    Respond ONLY with valid JSON.
    """)

    chain = prompt_template | llm

    result = chain.invoke({
        "topic": topic,
        "n_slides": n_slides,
        "instructions": instructions
    })

    content = result.content

    # Parse JSON output safely
    try:
        return json.loads(content)
    except Exception:
        return content


def build_presentation(slides: list[dict], template_path: str, output_path: str):
    """Populate slides into a PowerPoint using the company template."""
    from pptx import Presentation
    from pptx.util import Pt

    prs = Presentation(template_path)
    front_slide    = 0
    title_layout   = 2
    content_layout = 3

    def get_text_shapes(slide):
        return [s for s in slide.shapes if s.has_text_frame]

    prs.slides.add_slide(prs.slide_layouts[front_slide])

    # --- Cover slide ---
    cover = prs.slides.add_slide(prs.slide_layouts[title_layout])
    text_shapes = get_text_shapes(cover)
    if text_shapes:
        text_shapes[0].text = slides[0]["title"]
    else:
        print("⚠️ No text shapes found on cover layout")

    # Add speaker notes
    if "notes" in slides[0]:
        notes_frame = cover.notes_slide.notes_text_frame
        notes_frame.text = slides[0]["notes"]

    # --- Content slides ---
    for s in slides[1:]:
        slide = prs.slides.add_slide(prs.slide_layouts[content_layout])
        text_shapes = get_text_shapes(slide)

        if len(text_shapes) < 2:
            print(f"⚠️ Expected 2 text boxes, found {len(text_shapes)} on slide '{s['title']}'")
            continue

        # First box → title
        text_shapes[0].text = s["title"]

        # Second box → bullet content
        body = text_shapes[1].text_frame
        body.clear()
        for b in s["bullets"]:
            p = body.add_paragraph()
            p.text = b
            p.level = 0
            p.font.size = Pt(18)

        # Speaker notes
        if "notes" in s:
            notes_frame = slide.notes_slide.notes_text_frame
            notes_frame.text = s["notes"]

    prs.save(output_path)
    print(f"✅ Presentation saved to: {output_path}")

def main():
    parser = argparse.ArgumentParser(description="Generate PowerPoint slides with OpenAI.")
    parser.add_argument("--topic",    required=True, help="Presentation topic, e.g., 'DataCamp Overview'")
    parser.add_argument("--instructions",type=str,default="Make it professional and suitable for an internal company presentation.",help="Additional instructions for slide content")
    parser.add_argument("--slides", type=int, default=8, help="Number of slides to generate")
    parser.add_argument("--template", required=True, help="Path to company PowerPoint template")
    args = parser.parse_args()

    slides = generate_slide_outline(args.topic, args.slides, args.instructions)
    print(slides)
    build_presentation(slides['slides'], args.template, OUTPUT_FILE)


if __name__ == "__main__":
    main()
