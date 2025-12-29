#!/usr/bin/env python3
"""
PPTX Translation Tool - Translate PowerPoint text to Simplified Chinese
"""
from zipfile import ZipFile
import re
import boto3
import logging
from html import escape

DEBUG = False

bedrock_client = boto3.client("bedrock-runtime", region_name="us-west-2")

def setup_logging():
    """Configure logging"""
    level = logging.DEBUG if DEBUG else logging.INFO
    logging.basicConfig(format="%(message)s", level=level)
    # Suppress boto3 credentials info messages
    logging.getLogger("botocore.credentials").setLevel(logging.WARNING)


def call_llm(prompt):
    """Call Bedrock Claude Sonnet 4.5"""
    model_id = "us.anthropic.claude-sonnet-4-20250514-v1:0"
    messages = [{"role": "user", "content": [{"text": prompt}]}]

    response = bedrock_client.converse(modelId=model_id, messages=messages)
    return response["output"]["message"]["content"][0]["text"]


def extract_texts(xml_content):
    """Extract all <a:t> text using regex to avoid XML parsing issues"""
    pattern = r"<a:t>(.*?)</a:t>"
    texts = [match.group(1) for match in re.finditer(pattern, xml_content, re.DOTALL)]
    return xml_content, texts


def translate_texts(texts):
    """Batch translate texts"""
    if not texts:
        return []

    # Add ID to each tag for structure preservation
    xml_texts = "\n".join(f'<a:t id="{i}">{text}</a:t>' for i, text in enumerate(texts))

    prompt = (
        "You are a professional translator. Translate the text inside <a:t> tags to Simplified Chinese.\n\n"
        "CRITICAL RULES:\n"
        "1. Return EXACTLY the same number of <a:t> tags with matching id attributes\n"
        "2. Do NOT merge, split, or reorder tags\n"
        "3. Preserve formatting, line breaks, and whitespace within tags\n"
        "4. Keep unchanged: company names, product names, technical terms, proper nouns, numbers, URLs\n"
        "5. If text is already in Simplified Chinese or is empty, keep it as is\n\n"
        "EXAMPLE:\n"
        'Input:\n<a:t id="0">Hello World</a:t>\n<a:t id="1">AWS Lambda</a:t>\n'
        'Output:\n<a:t id="0">你好世界</a:t>\n<a:t id="1">AWS Lambda</a:t>\n\n'
        "INPUT:\n"
        f"{xml_texts}\n\n"
        "OUTPUT (XML only, no explanations):"
    )

    logging.debug("Prompt:\n%s", prompt)
    result = call_llm(prompt)
    logging.debug("LLM Response:\n%s", result)

    # Extract translated texts by ID to ensure correct mapping
    translated = []
    for i in range(len(texts)):
        # Match tag with specific id
        pattern = rf'<a:t id="{i}">(.*?)</a:t>'
        match = re.search(pattern, result, re.DOTALL)
        if match:
            translated.append(match.group(1))
        else:
            logging.warning("Missing translation for id=%d, using original text", i)
            translated.append(texts[i])

    if len(translated) != len(texts):
        logging.warning("Expected %d texts, got %d", len(texts), len(translated))
        logging.warning("%s\n%s\n%s", xml_texts, "#" * 80, result)
        return texts

    return translated


def replace_texts(xml_content, translated_texts):
    """Replace <a:t> text content in order"""
    counter = 0

    def replacer(match):
        nonlocal counter
        if counter < len(translated_texts):
            replacement = escape(translated_texts[counter])
            counter += 1
            return f"<a:t>{replacement}</a:t>"
        return match.group(0)

    pattern = r"<a:t>(.*?)</a:t>"
    result = re.sub(pattern, replacer, xml_content, flags=re.DOTALL)
    return result.encode("utf-8")


def translate_pptx(input_path, output_path):
    """Translate entire PPTX file"""
    logging.info("Processing: %s", input_path)

    with ZipFile(input_path, "r") as zip_in, ZipFile(output_path, "w") as zip_out:
        for item in zip_in.infolist():
            file_name = item.filename
            data = zip_in.read(file_name)

            # Process slides and notes
            if file_name.endswith(".xml") and (
                file_name.startswith("ppt/slides/slide")
                or file_name.startswith("ppt/notesSlides/notesSlide")
            ):
                logging.info("  Translating %s...", file_name.rsplit("/", 1)[-1][:-4])
                xml_content = data.decode("utf-8")
                _, texts = extract_texts(xml_content)
                logging.debug("    Found %d text elements", len(texts))

                if texts:
                    translated = translate_texts(texts)
                    data = replace_texts(xml_content, translated)

            zip_out.writestr(item, data)

    logging.info("Done! Output: %s", output_path)


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 3:
        print("Usage: python3 translate_pptx.py input.pptx output.pptx [--debug]")
        sys.exit(1)

    if "--debug" in sys.argv:
        DEBUG = True
        sys.argv.remove("--debug")

    setup_logging()
    translate_pptx(sys.argv[1], sys.argv[2])
