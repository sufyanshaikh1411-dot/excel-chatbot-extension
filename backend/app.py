from pathlib import Path
from flask import Flask, request, jsonify
from flask_cors import CORS
from openpyxl import load_workbook
import re

app = Flask(__name__)
CORS(app)

EXCEL_PATH = Path(__file__).resolve().parent.parent / "methodology.xlsx"
knowledge = []


# ---------- Helpers ----------

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


# ---------- Load Excel ----------

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

            # ✅ Month bucket detection (Column A or B, text or hyperlink)
            for cell in row[:2]:
                raw = cell.value
                if raw is None and cell.hyperlink:
                    raw = getattr(cell.hyperlink, "display", None)

                if not raw:
                    continue

                text = normalize_spaces(raw)
                if looks_like_month_group(text):
                    current_group_text = text
                    current_group_link = (
                        cell.hyperlink.target
                        if cell.hyperlink and cell.hyperlink.target
                        else None
                    )
                    current_group_month, current_group_year = normalize_month_and_year(text)
                    break

            row_texts = []
            row_links = []

            for cell in row:
                if cell.value:
                    row_texts.append(normalize_spaces(cell.value))
                if cell.hyperlink and cell.hyperlink.target:
                    row_links.append(cell.hyperlink.target)

            if not row_texts:
                continue

            knowledge.append({
                "sheet": sheet.title,
                "text": " | ".join(row_texts),
                "links": row_links,
                "month_group": current_group_text,
                "month_link": current_group_link,
                "month": current_group_month,
                "year": current_group_year,
            })


load_excel()


# ---------- Formatting ----------

def unique_items(items):
    seen = set()
    result = []
    for i in items:
        key = i["sheet"] + "||" + (i.get("month_group") or "") + "||" + i["text"]
        if key not in seen:
            seen.add(key)
            result.append(i)
    return result


def format_results(items, limit=5):
    if not items:
        return "No relevant data found."

    lines = []

    first = items[0]
    if first.get("month_group"):
        lines.append(first["month_group"])
        if first.get("month_link"):
            lines.append(first["month_link"])
        lines.append("-" * 40)

    for item in items[:limit]:
        lines.append(item["text"])
        for link in item.get("links", []):
            if link != item.get("month_link"):
                lines.append(link)
        lines.append("")

    return "\n".join(lines).strip()


# ---------- Search Logic (SIMPLIFIED) ----------

def search_answer(question, selected_sheet):
    q = question.lower().strip()
    q_month, q_year = normalize_month_and_year(q)

    sheet_items = [
        i for i in knowledge
        if i["sheet"].lower() == selected_sheet.lower()
    ]

    if not sheet_items:
        return f"No data found for sheet '{selected_sheet}'."

    # ✅ 1. Try month bucket (if user typed month/year)
    if q_month and q_year:
        bucket = [
            i for i in sheet_items
            if i["month"] == q_month and i["year"] == q_year
        ]
        bucket = unique_items(bucket)
        if bucket:
            return format_results(bucket)

    # ✅ 2. Always fallback to keyword search
    words = q.split()
    scored = []

    for item in sheet_items:
        score = sum(1 for w in words if len(w) > 2 and w in item["text"].lower())
        if score > 0:
            scored.append((score, item))

    if scored:
        scored.sort(key=lambda x: x[0], reverse=True)
        return format_results(unique_items([i for _, i in scored]))

    return f"No relevant data found in sheet '{selected_sheet}'."


# ---------- API ----------

@app.route("/chat", methods=["POST"])
def chat():
    data = request.get_json(silent=True) or {}
    question = str(data.get("question", "")).strip()
    sheet = str(data.get("sheet", "")).strip()

    if not question:
        return jsonify({"answer": "Please enter a question."})
    if not sheet:
        return jsonify({"answer": "Please select a sheet."})

    return jsonify({"answer": search_answer(question, sheet)})


@app.route("/", methods=["GET"])
def home():
    return "Excel chatbot backend is running."


if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
