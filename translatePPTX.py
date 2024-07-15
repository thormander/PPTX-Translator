import os
import json
import argparse
import requests
from pptx import Presentation
from tqdm import tqdm
from google.auth import jwt
from google.auth.transport.requests import Request

# Set the Google Cloud API key
API_KEY = "Google_API.json"
os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = API_KEY

def get_access_token():
    with open(API_KEY, 'r') as f:
        credentials = json.load(f)
    credentials_info = jwt.Credentials.from_service_account_info(credentials)
    scoped_credentials = credentials_info.with_scopes(['https://www.googleapis.com/auth/cloud-platform'])
    auth_req = Request()
    scoped_credentials.refresh(auth_req)
    return scoped_credentials.token

def translate_text(text, access_token, target_language):
    url = "https://translation.googleapis.com/v3/projects/your_project_id:translateText"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    body = {
        "targetLanguageCode": target_language,
        "contents": [text],
        "mimeType": "text/plain"
    }
    response = requests.post(url, headers=headers, json=body)
    if response.status_code == 200:
        return response.json()['translations'][0]['translatedText']
    else:
        print(f"Error translating text: {response.status_code} {response.text}")
        return text

def translate_shape_text(shape, access_token, target_language):
    if not hasattr(shape, "text_frame") or not shape.text_frame:
        return

    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            translated_text = translate_text(run.text, access_token, target_language)
            run.text = translated_text

def process_presentation(input_file, target_language):
    print(f"Opening {input_file}")
    try:
        input_ppt = Presentation(input_file)
    except Exception as e:
        print(f"Error opening file {input_file}: {e}")
        return

    output_ppt = Presentation()
    access_token = get_access_token()
    slide_count = len(input_ppt.slides)
    
    with tqdm(total=slide_count, desc="Translating", unit="slide") as pbar:
        for i, slide in enumerate(input_ppt.slides):
            new_slide_layout = output_ppt.slide_layouts[5]
            new_slide = output_ppt.slides.add_slide(new_slide_layout)
            for shape in slide.shapes:
                if hasattr(shape, "text_frame") and shape.text_frame:
                    try:
                        translate_shape_text(shape, access_token, target_language)
                        new_shape = new_slide.shapes.add_textbox(shape.left, shape.top, shape.width, shape.height)
                        new_text_frame = new_shape.text_frame
                        for paragraph in shape.text_frame.paragraphs:
                            new_paragraph = new_text_frame.add_paragraph()
                            for run in paragraph.runs:
                                new_run = new_paragraph.add_run()
                                new_run.text = run.text
                                new_run.font.bold = run.font.bold
                                new_run.font.italic = run.font.italic
                                new_run.font.size = run.font.size
                                new_run.font.color.rgb = run.font.color.rgb
                    except Exception as e:
                        print(f"Error processing shape on slide {i}: {e}")
            pbar.update(1)

    output_file = f"{target_language}_{input_file}"
    try:
        output_ppt.save(output_file)
        print(f"\nSaved as {output_file}")
    except Exception as e:
        print(f"Error saving file {output_file}: {e}")

def main():
    parser = argparse.ArgumentParser(description="Translate a PowerPoint presentation. Usage: python3 translatePPTX.py <input_pptx_file> <target_language>")
    parser.add_argument("input_file", help="Path to the input PowerPoint file")
    parser.add_argument("target_language", help="Target language for translation (e.g., 'en' for English, 'fr' for French)")
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