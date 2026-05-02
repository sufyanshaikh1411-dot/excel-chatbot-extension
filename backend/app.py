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
    if not text:
        return None, None

    text = str(text).lower()
    text = text.replace(",", " ").replace("-", " ")
    text = re.sub(r"\s+", " ", text).strip()

    month_map = {
        "jan": "january", "january": "january",
        "feb": "february", "february": "february",
        "mar": "march", "march": "march",
        "apr": "april", "april": "april",
        "may": "may",
        "jun": "june", "june": "june",
        "jul": "july", "july": "july",
        "aug": "august", "august": "august",
        "sep": "september", "sept": "september", "september": "september",
        "oct": "october", "october": "october",
        "nov": "november", "november": "november",
        "dec": "december", "december": "december",
    }

    parts = text.split()
    month = None
    year = None

    for p in parts:
        if p in month_map:
            month = month_map[p]
            break

    for p in parts:
        if re.fullmatch(r"20\d{2}", p):
            year = p
            break
        if re.fullmatch(r"\d{2}", p):
            year = f"20{p}"
            break

    return month, year


def looks_like_month_group(text):
    m, y = normalize_month_and_year(text)
    return bool(m and y)


def load_excel():
    global knowledge
    knowledge = []

    wb = load_workbook(EXCEL_PATH, data_only=True)

    for sheet in wb.worksheets:
        current_group_text = None
        current_group_link = None
        current_group_month = None
        current_group_year = None

        for row in sheet.iter_rows():

            # ✅ Detect linked or normal month headers in column A or B
            for cell in row[:2]:
                raw_value = cell.value

                if raw_value is None and cell.hyperlink:
                    raw_value = getattr(cell.hyperlink, "display", None)

                if not raw_value:
                    continue

                cell_text = normalize_spaces(raw_value)

                if looks_like_month_group(cell_text):
                    current_group_text = cell_text
                    current_group_link = (
                        cell.hyperlink.target
                        if cell.hyperlink and cell.hyperlink.target
                        else None
                    )
                    current_group_month, current_group_year = normalize_month_and_year(cell_text)
                    break

            row_values = []
            row_links = []

            for cell in row:
                if cell.value:
                    txt = normalize_spaces(cell.value)
                    if txt:
                        row_values.append(txt)

                if cell.hyperlink and cell.hyperlink.target:
                    row_links.append(cell.hyperlink.target)

            if not row_values:
                continue

            knowledge.append({
                "sheet": sheet.title,
                "text": " | ".join(row_values),
                "links": row_links,
                "month_group": current_group_text,
                "month_link": current_group_link,
                "month": current_group_month,
                "year": current_group_year,
            })


load_excel()


def unique_items(items):
    seen = set()
    out = []

    for i in items:
        key = (
            i["sheet"]
            + "||"
            + str(i.get("month_group") or "")
            + "||"
            + i["text"]
        )
        if key not in seen:
            seen.add(key)
            out.append(i)

    return out


def format_bucket_results(items):
    first = items[0]
    lines = []

    if first.get("month_group"):
        lines.append(first["month_group"])
        if first.get("month_link"):
            lines.append(first["month_link"])
        lines.append("-" * 40)

    for i in items:
        lines.append(i["text"])

        for l in i.get("links", []):
            if l != i.get("month_link"):
                lines.append(l)

        lines.append("")

    return "\n".join(lines).strip()


def score_keyword_match(text, words):
    text_words = set(re.findall(r"\b\w+\b", text.lower()))
    return sum(1 for w in words if len(w) > 2 and w in text_words)


def search_answer(question, sheet_name):
    q = question.lower().strip()
    q_month, q_year = normalize_month_and_year(q)

    sheet_items = [
        i for i in knowledge
        if i["sheet"].lower() == sheet_name.lower()
    ]

    if not sheet_items:
        return f"No data found for sheet '{sheet_name}'."

    # ✅ Month query
    if q_month and q_year:
        matches = [
            i for i in sheet_items
            if i["month"] == q_month and i["year"] == q_year
        ]

        matches = unique_items(matches)

        if matches:
            return format_bucket_results(matches)

        return f"No updates found for {q_month.title()} {q_year} in sheet '{sheet_name}'."

    # ✅ Keyword search
    words = q.split()
    ranked = []

    for i in sheet_items:
        score = score_keyword_match(i["text"], words)
        if score > 0:
            ranked.append((score, i))

    if ranked:
        ranked.sort(key=lambda x: x[0], reverse=True)
        return format_bucket_results(unique_items([i for _, i in ranked][:5]))

    return f"No relevant data found in sheet '{sheet_name}'."


@app.route("/chat", methods=["POST"])
def chat():
    data = request.get_json(silent=True) or {}
    q = str(data.get("question", "")).strip()
    sheet = str(data.get("sheet", "")).strip()

    if not q:
        return jsonify({"answer": "Please enter a question."})
    if not sheet:
        return jsonify({"answer": "Please select a sheet."})

    return jsonify({"answer": search_answer(q, sheet)})


@app.route("/", methods=["GET"])
def home():
    return "Excel chatbot backend is running."


if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
