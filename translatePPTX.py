import os
import argparse
import requests
from pptx import Presentation
from tqdm import tqdm

# Set your Google Cloud API key
API_KEY = "API KEY HERE"

def translate_text(text, target_language):
    url = f"https://translation.googleapis.com/language/translate/v2?key={API_KEY}"
    headers = {
        "Content-Type": "application/json"
    }
    body = {
        "q": text,
        "target": target_language,
        "format": "text"
    }
    response = requests.post(url, headers=headers, json=body)
    if response.status_code == 200:
        return response.json()['data']['translations'][0]['translatedText']
    else:
        print(f"Error translating text: {response.status_code} {response.text}")
        return text

def translate_shape_text(shape, target_language):
    if not hasattr(shape, "text_frame") or not shape.text_frame:
        return

    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            translated_text = translate_text(run.text, target_language)
            run.text = translated_text

def process_presentation(input_file, target_language):
    print(f"Opening {input_file}")
    try:
        input_ppt = Presentation(input_file)
    except Exception as e:
        print(f"Error opening file {input_file}: {e}")
        return

    slide_count = len(input_ppt.slides)
    
    with tqdm(total=slide_count, desc="Translating", unit="slide") as pbar:
        for i, slide in enumerate(input_ppt.slides):
            for shape in slide.shapes:
                if hasattr(shape, "text_frame") and shape.text_frame:
                    try:
                        translate_shape_text(shape, target_language)
                    except Exception as e:
                        print(f"Error processing shape on slide {i}: {e}")
            pbar.update(1)

    output_file = f"{target_language}_{input_file}"
    try:
        input_ppt.save(output_file)
        print(f"\nSaved as {output_file}")
    except Exception as e:
        print(f"Error saving file {output_file}: {e}")

def main():
    parser = argparse.ArgumentParser(description="Translate a PowerPoint presentation. Usage: python3 translatePPTX.py <input_pptx_file> <target_language>")
    parser.add_argument("input_file", help="Path to the input PowerPoint file")
    parser.add_argument("target_language", help="Target language for translation (ex: 'en' for English, 'fr' for French)")
    args = parser.parse_args()

    print("Example language syntax:")
    example_usages = [
        ("English", "en"),
        ("Spanish", "es"),
        ("French", "fr"),
        ("German", "de"),
        ("Italian", "it"),
        ("Portuguese", "pt"),
        ("Russian", "ru"),
        ("Chinese (Simplified)", "zh-CN"),
        ("Japanese", "ja"),
        ("Korean", "ko")
    ]

    for language, code in example_usages:
        print(f"  python translatePPTX.py <your_PPTX_file> {code}  # {language}")

    process_presentation(args.input_file, args.target_language)

if __name__ == "__main__":
    main()
