"""
LLM-based structured data extractor.
Uses a local Ollama model to analyze email content and extract structured fields:
Date, Brand, Cartons, Units, Adhoc, Comments, PO
"""

import json
import requests

OLLAMA_URL = "http://localhost:11434/api/generate"
DEFAULT_MODEL = "llama3.1:latest"

# The fields we want to extract from each email
TARGET_FIELDS = ["Date", "Brand", "Cartons", "Units", "Adhoc", "Comments", "PO"]

EXTRACTION_PROMPT = """You are a data extraction assistant. Analyze the email below and extract the following fields into JSON.

Fields to extract:
- Date: any relevant date mentioned (delivery date, PO date, etc.) in YYYY-MM-DD format if possible
- Brand: the brand name(s) mentioned
- Cartons: number of cartons (numeric only, or empty if not present)
- Units: number of units (numeric only, or empty if not present)
- Adhoc: any adhoc quantity or note about adhoc requests
- Comments: any relevant comments, notes, or special instructions
- PO: purchase order number(s)

Rules:
- Return ONLY a valid JSON object with exactly these keys: Date, Brand, Cartons, Units, Adhoc, Comments, PO
- If a field is not found, use an empty string ""
- Do not invent data. Only extract what is actually present.
- If multiple values exist for a field, join them with "; "
- Do not include any explanation, only the JSON object.

EMAIL SUBJECT: {subject}

EMAIL CONTENT:
{content}

JSON:"""


def check_ollama_available(model=DEFAULT_MODEL):
    """Check if Ollama is running and the model is available."""
    try:
        resp = requests.get("http://localhost:11434/api/tags", timeout=5)
        if resp.status_code != 200:
            return False, "Ollama server not responding"
        models = [m["name"] for m in resp.json().get("models", [])]
        if model not in models:
            return False, f"Model '{model}' not found. Available: {', '.join(models) or 'none'}"
        return True, "OK"
    except requests.exceptions.RequestException:
        return False, "Ollama is not running. Start it and try again."


def extract_fields_from_email(subject, content, model=DEFAULT_MODEL, timeout=120):
    """
    Send an email's content to Ollama and extract structured fields.
    Returns a dict with the TARGET_FIELDS keys.
    """
    # Empty result template
    empty_result = {field: "" for field in TARGET_FIELDS}

    if not content or not content.strip():
        return empty_result

    # Truncate very long content to keep the model fast and within context
    max_chars = 8000
    if len(content) > max_chars:
        content = content[:max_chars] + "\n...[truncated]"

    prompt = EXTRACTION_PROMPT.format(subject=subject or "", content=content)

    payload = {
        "model": model,
        "prompt": prompt,
        "stream": False,
        "format": "json",  # Ask Ollama to constrain output to valid JSON
        "options": {
            "temperature": 0,  # Deterministic extraction
            "num_predict": 512,
        },
    }

    try:
        resp = requests.post(OLLAMA_URL, json=payload, timeout=timeout)
        resp.raise_for_status()
        raw = resp.json().get("response", "").strip()

        # Parse the JSON response
        parsed = json.loads(raw)

        # Normalize: ensure all target fields exist and are strings
        result = {}
        for field in TARGET_FIELDS:
            val = parsed.get(field, "")
            if val is None:
                val = ""
            result[field] = str(val).strip()
        return result

    except json.JSONDecodeError:
        # Model returned non-JSON; try to salvage or return empty
        return empty_result
    except requests.exceptions.RequestException as e:
        raise ConnectionError(f"Ollama request failed: {e}")


def extract_batch(emails, model=DEFAULT_MODEL, progress_callback=None):
    """
    Extract structured fields from a list of email dicts.
    Each email dict should have 'Subject', 'Body', and optionally 'AttachmentText'.

    Args:
        emails: list of dicts with email data
        model: Ollama model name
        progress_callback: optional function(current, total) for progress reporting

    Returns:
        list of dicts, each combining original email metadata with extracted fields
    """
    results = []
    total = len(emails)

    for i, email in enumerate(emails):
        subject = email.get("Subject", "")
        body = email.get("Body", "") or ""
        attachment_text = email.get("AttachmentText", "") or ""

        # Combine body and attachment text
        combined = body
        if attachment_text:
            combined += "\n\n--- ATTACHMENTS ---\n" + attachment_text

        try:
            extracted = extract_fields_from_email(subject, combined, model=model)
        except ConnectionError:
            extracted = {field: "" for field in TARGET_FIELDS}

        # Merge original metadata with extracted fields
        row = {
            "ReceivedTime": email.get("ReceivedTime", ""),
            "Subject": subject,
            "Sender": email.get("Sender", ""),
            "Mailbox": email.get("Mailbox", ""),
            "FolderName": email.get("FolderName", ""),
        }
        row.update(extracted)
        results.append(row)

        if progress_callback:
            progress_callback(i + 1, total)

    return results


if __name__ == "__main__":
    ok, msg = check_ollama_available()
    print(f"Ollama check: {msg}")
    if ok:
        sample = extract_fields_from_email(
            "PO 12345 - Nike shipment",
            "Hi team, please find the daily flash. Brand: Nike, 50 cartons, 1200 units. "
            "PO number is 12345. Delivery expected on 2026-08-25. Adhoc request for 10 extra cartons. "
            "Comments: urgent, prioritize this shipment.",
        )
        print(json.dumps(sample, indent=2))
