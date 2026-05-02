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
        current_month_link = None

        for row in sheet.iter_rows():
            row_values = []
            row_links = []

            for idx, cell in enumerate(row):
                cell_value = cell.value

                if cell_value is not None:
                    text = str(cell_value).strip()
                    if text:
                        row_values.append(text)

                        # first column month-year group capture
                        if idx == 0:
                            month_year_pattern = r"(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|sept|oct|nov|dec)[,\s-]*20\d{2}"
                            if re.search(month_year_pattern, text.lower()):
                                current_month_year = text
                                current_month_link = cell.hyperlink.target if cell.hyperlink and cell.hyperlink.target else None

                if cell.hyperlink and cell.hyperlink.target:
                    row_links.append(cell.hyperlink.target)

            if row_values or row_links:
                joined_text = " | ".join(row_values)

                knowledge.append({
                    "sheet": sheet.title,
                    "text": joined_text,
                    "links": row_links,
                    "group": current_month_year.lower() if current_month_year else "",
                    "month": current_month_year,
                    "month_link": current_month_link
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


def unique_items_only(items):
    unique = []
    seen = set()

    for item in items:
        key = (
            item["sheet"]
            + "||"
            + item["text"]
            + "||"
            + (item.get("month") or "")
            + "||"
            + "||".join(item.get("links", []))
        )
        if key not in seen:
            seen.add(key)
            unique.append(item)

    return unique


def format_results(items, limit=None):
    if not items:
        return "No relevant data found."

    output = []
    count = 0

    for item in items:
        if limit is not None and count >= limit:
            break

        parts = []

        # Month-Year should come FIRST / UPPER
        if item.get("month"):
            if item.get("month_link"):
                parts.append(f"{item['month']}\n{item['month_link']}")
            else:
                parts.append(item["month"])

        # Then matched row text
        parts.append(item["text"])

        # Then any row-level links except duplicated month link
        extra_links = []
        for link in item.get("links", []):
            if link != item.get("month_link"):
                extra_links.append(link)

        if extra_links:
            parts.extend(extra_links)

        output.append("\n".join(parts))
        count += 1

    return "\n\n".join(output)


def search_answer(question, selected_sheet):
    q = question.lower().strip()
    month, year, month_map = normalize_month_query(q)

    cleaned_query = q

    if year:
        cleaned_query = re.sub(rf"\b{re.escape(year)}\b", " ", cleaned_query)

    for key in month_map.keys():
        cleaned_query = re.sub(rf"\b{re.escape(key)}\b", " ", cleaned_query)

    keywords = [word.strip() for word in cleaned_query.split() if word.strip()]

    filtered_knowledge = [
        item for item in knowledge
        if item["sheet"].strip().lower() == selected_sheet.strip().lower()
    ]

    month_year_matches = []
    keyword_matches = []
    fallback_matches = []

    for item in filtered_knowledge:
        text_lower = item["text"].lower()
        group_lower = item.get("group", "").lower()

        # exact month-year query -> all rows from that group
        if month and year:
            if month in group_lower and year in group_lower:
                month_year_matches.append(item)
                continue

            if month in text_lower and year in text_lower:
                month_year_matches.append(item)
                continue

        # keyword search
        keyword_score = 0
        for word in keywords:
            if word in text_lower:
                keyword_score += 1

        if keyword_score > 0:
            keyword_matches.append((keyword_score, item))
            continue

        # generic fallback
        fallback_score = 0
        for word in q.split():
            if word in text_lower:
                fallback_score += 1

        if fallback_score > 0:
            fallback_matches.append((fallback_score, item))

    if month and year and month_year_matches:
        items = unique_items_only(month_year_matches)
        return format_results(items, limit=None)

    if keyword_matches:
        keyword_matches.sort(key=lambda x: x[0], reverse=True)
        items = [item for _, item in keyword_matches]
        items = unique_items_only(items)
        return format_results(items, limit=20)

    if fallback_matches:
        fallback_matches.sort(key=lambda x: x[0], reverse=True)
        items = [item for _, item in fallback_matches]
        items = unique_items_only(items)
        return format_results(items, limit=5)

    return f"No relevant data found in sheet '{selected_sheet}'."


@app.route("/chat", methods=["POST"])
def chat():
    data = request.get_json(silent=True) or {}
    question = data.get("question", "").strip()
    selected_sheet = data.get("sheet", "").strip()

    if not question:
        return jsonify({"answer": "Please enter a question."})

    if not selected_sheet:
        return jsonify({"answer": "Please select a sheet."})

    answer = search_answer(question, selected_sheet)
    return jsonify({"answer": answer})


@app.route("/", methods=["GET"])
def home():
    return "Excel chatbot backend is running."


if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
