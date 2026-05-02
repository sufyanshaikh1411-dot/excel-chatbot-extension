from pathlib import Path
from flask import Flask, request, jsonify
from flask_cors import CORS
from openpyxl import load_workbook
import re

app = Flask(__name__)
CORS(app)

EXCEL_PATH = Path(__file__).resolve().parent.parent / "methodology.xlsx"

knowledge = []


def load_excel():
    global knowledge
    knowledge = []

    wb = load_workbook(EXCEL_PATH, data_only=True)

    for sheet in wb.worksheets:
        current_month_year = None

        for row in sheet.iter_rows():
            row_values = []
            row_links = []

            for idx, cell in enumerate(row):
                cell_value = cell.value

                if cell_value is not None:
                    text = str(cell_value).strip()
                    if text:
                        row_values.append(text)

                        # First meaningful column often carries month/year group
                        if idx == 0 and re.search(r"(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|sept|oct|nov|dec)[,\s-]*20\d{2}", text.lower()):
                            current_month_year = text

                if cell.hyperlink and cell.hyperlink.target:
                    row_links.append(cell.hyperlink.target)

            if row_values or row_links:
                joined_text = " | ".join(row_values)

                knowledge.append({
                    "sheet": sheet.title,
                    "text": joined_text,
                    "links": row_links,
                    "group": current_month_year.lower() if current_month_year else ""
                })


load_excel()


def normalize_month_query(text):
    text = text.lower().strip()

    month_map = {
        "jan": "january",
        "january": "january",
        "feb": "february",
        "february": "february",
        "mar": "march",
        "march": "march",
        "apr": "april",
        "april": "april",
        "may": "may",
        "jun": "june",
        "june": "june",
        "jul": "july",
        "july": "july",
        "aug": "august",
        "august": "august",
        "sep": "september",
        "sept": "september",
        "september": "september",
        "oct": "october",
        "october": "october",
        "nov": "november",
        "november": "november",
        "dec": "december",
        "december": "december"
    }

    detected_month = None
    for key, value in month_map.items():
        if re.search(rf"\b{re.escape(key)}\b", text):
            detected_month = value
            break

    year_match = re.search(r"\b(20\d{2})\b", text)
    detected_year = year_match.group(1) if year_match else None

    return detected_month, detected_year, month_map


def format_results(items, limit=None):
    if not items:
        return "No relevant data found."

    output = []
    count = 0

    for item in items:
        if limit is not None and count >= limit:
            break

        line = item["text"]

        if item.get("links"):
            for link in item["links"]:
                line += f"\n{link}"

        output.append(line)
        count += 1

    return "\n\n".join(output)


def search_answer(question):
    q = question.lower().strip()
    month, year, month_map = normalize_month_query(q)

    cleaned_query = q

    if year:
        cleaned_query = re.sub(rf"\b{re.escape(year)}\b", " ", cleaned_query)

    for key in month_map.keys():
        cleaned_query = re.sub(rf"\b{re.escape(key)}\b", " ", cleaned_query)

    keywords = [word.strip() for word in cleaned_query.split() if word.strip()]

    month_year_matches = []
    keyword_matches = []
    fallback_matches = []

    for item in knowledge:
        text_lower = item["text"].lower()
        group_lower = item.get("group", "").lower()

        # MODE 1: exact month + year group match
        if month and year:
            if month in group_lower and year in group_lower:
                month_year_matches.append(item)
                continue

            # fallback to text if group not captured well
            if month in text_lower and year in text_lower:
                month_year_matches.append(item)
                continue

        # MODE 2: keyword-specific search
        keyword_score = 0
        for word in keywords:
            if word in text_lower:
                keyword_score += 1

        if keyword_score > 0:
            keyword_matches.append((keyword_score, item))
            continue

        # MODE 3: generic fallback search
        fallback_score = 0
        for word in q.split():
            if word in text_lower:
                fallback_score += 1

        if fallback_score > 0:
            fallback_matches.append((fallback_score, item))

    # Return ALL rows for month + year
    if month and year and month_year_matches:
        unique_items = []
        seen = set()

        for item in month_year_matches:
            key = item["text"] + "||" + "||".join(item.get("links", []))
            if key not in seen:
                seen.add(key)
                unique_items.append(item)

        return format_results(unique_items, limit=None)

    # Return related keyword rows, up to 20
    if keyword_matches:
        keyword_matches.sort(key=lambda x: x[0], reverse=True)

        unique_items = []
        seen = set()

        for _, item in keyword_matches:
            key = item["text"] + "||" + "||".join(item.get("links", []))
            if key not in seen:
                seen.add(key)
                unique_items.append(item)

        return format_results(unique_items, limit=20)

    # Generic fallback = top 5
    if fallback_matches:
        fallback_matches.sort(key=lambda x: x[0], reverse=True)

        unique_items = []
        seen = set()

        for _, item in fallback_matches:
            key = item["text"] + "||" + "||".join(item.get("links", []))
            if key not in seen:
                seen.add(key)
                unique_items.append(item)

        return format_results(unique_items, limit=5)

    return "No relevant data found."


@app.route("/chat", methods=["POST"])
def chat():
    data = request.get_json(silent=True) or {}
    question = data.get("question", "").strip()

    if not question:
        return jsonify({"answer": "Please enter a question."})

    answer = search_answer(question)
    return jsonify({"answer": answer})


@app.route("/", methods=["GET"])
def home():
    return "Excel chatbot backend is running."


if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
