# PPTX Translator

PPTX Translator is a Python script to translate text in PowerPoint presentations to a specified language using the Google Cloud Translation API.

The goal of this is to **retain original formatting** of the powerpoint and only translate the text

## Requirements

You need to create an account at [Google Cloud Console](https://cloud.google.com/cloud-console) --> they have a free trial for 90 days
- Enable Translate API
- Create an API key
- Create a .env file at same location as script and add 'GOOGLE_API_KEY=YOUR_KEY_HERE'
  - ![image](https://github.com/user-attachments/assets/0399b3da-e63f-4e47-a397-8b8400e37a5f)


*For our case, we use v2, so we do not need to mess around with OAUTH, we can just use an API KEY*



## Usage
```console
python3 translatePPTX.py [-h] [--list-langs] <PPTX file you want to translate> <target language>
```
Arguments within brackets [...] are optional
- [-h] is to for help
- [--list-langs] is to list all compatiable languages

Example Usage for accessing list of available languages: 
```console
python3 translatePPTX.py --list-langs
```

## Packages

You need to install the following Python packages:

```sh
pip install requests python-pptx tqdm
pip install python-dotenv
```

