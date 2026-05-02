from pathlib import Path
from flask import Flask, request, jsonify
from flask_cors import CORS
from openpyxl import load_workbook
import re

app = Flask(__name__)
CORS(app)

EXCEL_PATH = Path(__file__).resolve().parent.parent / "methodology.xlsx"

knowledge = []


def normalize_spaces(text):
    return re.sub(r"\s+", " ", str(text).strip())


def normalize_month_and_year(text):
    """
    Supports:
    Jan 26
    Jan-26
    Jan,26
    January 2026
    March,2026
    """
    if not text:
        return None, None

    text = str(text).lower().strip()
    text = text.replace(",", " ").replace("-", " ")
    text = re.sub(r"\s+", " ", text)

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
        "december": "december",
    }

    parts = text.split()

    detected_month = None
    detected_year = None

    for part in parts:
        if part in month_map:
            detected_month = month_map[part]
            break

    for part in parts:
        if re.fullmatch(r"20\d{2}", part):
            detected_year = part
            break
        if re.fullmatch(r"\d{2}", part):
            detected_year = f"20{part}"
            break

    return detected_month, detected_year


def looks_like_month_group(text):
    month, year = normalize_month_and_year(text)
    return bool(month and year)


def load_excel():
    """
    Reads workbook and links Column A month/year header + its hyperlink
    to following B/C rows until next month/year header appears.
    """
    global knowledge
    knowledge = []

    wb = load_workbook(EXCEL_PATH, data_only=True)

    for sheet in wb.worksheets:
        current_group_text = None
        current_group_link = None
        current_group_month = None
        current_group_year = None

        for row in sheet.iter_rows():
            # Column A handling
            col_a_text = ""
            col_a_link = None

            if len(row) >= 1:
                a_cell = row[0]

                if a_cell.value is not None:
                    col_a_text = normalize_spaces(a_cell.value)

                if a_cell.hyperlink and a_cell.hyperlink.target:
                    col_a_link = a_cell.hyperlink.target

            # Update current month/year group if column A has a month-year header
            if looks_like_month_group(col_a_text):
                current_group_text = col_a_text
                current_group_link = col_a_link
                current_group_month, current_group_year = normalize_month_and_year(col_a_text)

            row_values = []
            row_links = []

            for cell in row:
                if cell.value is not None:
                    text = normalize_spaces(cell.value)
                    if text:
                        row_values.append(text)

                if cell.hyperlink and cell.hyperlink.target:
                    row_links.append(cell.hyperlink.target)

            if not row_values and not row_links:
                continue

            joined_text = " | ".join(row_values)

            knowledge.append({
                "sheet": sheet.title,
                "text": joined_text,
                "links": row_links,
                "month_group": current_group_text,
                "month_link": current_group_link,
                "month": current_group_month,
                "year": current_group_year,
            })


load_excel()


def unique_items(items):
    unique = []
    seen = set()

    for item in items:
        key = (
            item["sheet"]
            + "||"
            + str(item.get("month_group") or "")
            + "||"
            + item["text"]
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

        block = []

        # Month header first
        if item.get("month_group"):
            if item.get("month_link"):
                block.append(f"{item['month_group']}\n{item['month_link']}")
            else:
                block.append(item["month_group"])

        # Main row text
        block.append(item["text"])

        # Row-level links except duplicate month link
        extra_links = []
        for link in item.get("links", []):
            if link != item.get("month_link"):
                extra_links.append(link)

        if extra_links:
            block.extend(extra_links)

        output.append("\n".join(block))
        count += 1

    return "\n\n".join(output)


def score_keyword_match(text, keywords):
    score = 0
    text_lower = text.lower()

    for word in keywords:
        if word in text_lower:
            score += 1

    return score


def search_answer(question, selected_sheet):
    q = question.lower().strip()

    query_month, query_year = normalize_month_and_year(q)

    # Remove month/year tokens for keyword-only search
    cleaned_query = q
    month_words = [
        "jan", "january", "feb", "february", "mar", "march", "apr", "april",
        "may", "jun", "june", "jul", "july", "aug", "august", "sep", "sept",
        "september", "oct", "october", "nov", "november", "dec", "december"
    ]

    for word in month_words:
        cleaned_query = re.sub(rf"\b{re.escape(word)}\b", " ", cleaned_query)

    cleaned_query = re.sub(r"\b20\d{2}\b", " ", cleaned_query)
    cleaned_query = re.sub(r"\b\d{2}\b", " ", cleaned_query)
    cleaned_query = re.sub(r"\s+", " ", cleaned_query).strip()

    keywords = [w.strip() for w in cleaned_query.split() if w.strip()]

    filtered = [
        item for item in knowledge
        if item["sheet"].strip().lower() == selected_sheet.strip().lower()
    ]

    if not filtered:
        return f"No data found for sheet '{selected_sheet}'."

    # CASE 1: month + year query => return ALL updates from that month group
    if query_month and query_year:
        month_matches = []

        for item in filtered:
            item_month = item.get("month")
            item_year = item.get("year")

            if item_month == query_month and item_year == query_year:
                month_matches.append(item)

        month_matches = unique_items(month_matches)

        if month_matches:
            return format_results(month_matches, limit=None)

        return f"No updates found for {query_month.title()} {query_year} in sheet '{selected_sheet}'."

    # CASE 2: keyword search => top 5
    scored = []

    for item in filtered:
        score = score_keyword_match(item["text"], keywords)

        if score > 0:
            scored.append((score, item))

    if scored:
        scored.sort(key=lambda x: x[0], reverse=True)
        ranked_items = [item for _, item in scored]
        ranked_items = unique_items(ranked_items)
        return format_results(ranked_items, limit=5)

    # CASE 3: generic fallback => top 5
    fallback = []
    query_words = [w.strip() for w in q.split() if w.strip()]

    for item in filtered:
        score = score_keyword_match(item["text"], query_words)
        if score > 0:
            fallback.append((score, item))

    if fallback:
        fallback.sort(key=lambda x: x[0], reverse=True)
        fallback_items = [item for _, item in fallback]
        fallback_items = unique_items(fallback_items)
        return format_results(fallback_items, limit=5)

    return f"No relevant data found in sheet '{selected_sheet}'."


@app.route("/chat", methods=["POST"])
def chat():
    data = request.get_json(silent=True) or {}

    question = str(data.get("question", "")).strip()
    selected_sheet = str(data.get("sheet", "")).strip()

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
