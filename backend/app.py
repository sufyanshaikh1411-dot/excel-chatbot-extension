from pathlib import Path
from flask import Flask, request, jsonify
from flask_cors import CORS
from openpyxl import load_workbook
from datetime import datetime, date
import re

app = Flask(__name__)
CORS(app)

EXCEL_PATH = Path(__file__).resolve().parent.parent / "methodology.xlsx"
knowledge = []

MONTH_ORDER = {
    "january": 1, "february": 2, "march": 3, "april": 4,
    "may": 5, "june": 6, "july": 7, "august": 8,
    "september": 9, "october": 10, "november": 11, "december": 12
}

# ---------------- Helpers ----------------

def normalize_spaces(text):
    return re.sub(r"\s+", " ", str(text).strip())


def normalize_month_and_year(text):
    if not text:
        return None, None

    text = str(text).lower().replace(",", " ").replace("-", " ")
    text = re.sub(r"\s+", " ", text).strip()

    months = {
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

    month = year = None
    for p in text.split():
        if p in months:
            month = months[p]
        if re.fullmatch(r"20\d{2}", p):
            year = p
        elif re.fullmatch(r"\d{2}", p):
            year = f"20{p}"

    return month, year


# ---------------- Load Excel ----------------

def load_excel():
    knowledge.clear()
    wb = load_workbook(EXCEL_PATH, data_only=True)

    for sheet in wb.worksheets:
        cur_text = cur_link = cur_month = cur_year = None

        for row in sheet.iter_rows():
            # month bucket (A or B)
            for cell in row[:2]:
                raw = cell.value

                if isinstance(raw, (datetime, date)):
                    cur_month = raw.strftime("%B").lower()
                    cur_year = str(raw.year)
                    cur_text = raw.strftime("%b-%y")
                    cur_link = cell.hyperlink.target if cell.hyperlink else None
                    break

                if raw is None and cell.hyperlink:
                    raw = getattr(cell.hyperlink, "display", None)

                if raw:
                    m, y = normalize_month_and_year(raw)
                    if m and y:
                        cur_month, cur_year = m, y
                        cur_text = normalize_spaces(raw)
                        cur_link = cell.hyperlink.target if cell.hyperlink else None
                        break

            texts, links = [], []
            for cell in row:
                if cell.value:
                    texts.append(normalize_spaces(cell.value))
                if cell.hyperlink and cell.hyperlink.target:
                    links.append(cell.hyperlink.target)

            if texts:
                knowledge.append({
                    "sheet": sheet.title,
                    "text": " | ".join(texts),
                    "links": links,
                    "month_group": cur_text,
                    "month_link": cur_link,
                    "month": cur_month,
                    "year": cur_year,
                })


load_excel()

# ---------------- Formatting ----------------

def format_results(items):
    output, seen = [], set()
    for item in items:
        key = (item["month"], item["year"])
        if key not in seen:
            seen.add(key)
            output.append(item["month_group"])
            if item["month_link"]:
                output.append(item["month_link"])
            output.append("-" * 40)

        if item["text"]:
            output.append(item["text"])
            for l in item["links"]:
                if l != item["month_link"]:
                    output.append(l)
            output.append("")

    return "\n".join(output).strip()


# ---------------- Search Logic (FIXED) ----------------

def search_answer(question, selected_sheet):
    q = question.lower().strip()
    sheet_items = [i for i in knowledge if i["sheet"].lower() == selected_sheet.lower()]

    words = [w for w in q.split() if len(w) > 2]
    matches = []

    # ✅ keyword exact matches only
    for item in sheet_items:
        if any(w in item["text"].lower() for w in words):
            matches.append(item)

    if not matches:
        return f"No relevant data found in sheet '{selected_sheet}'."

    # ✅ get matched month buckets only
    unique_months = {}
    for i in matches:
        k = (i["year"], MONTH_ORDER.get(i["month"]))
        unique_months[k] = i

    sorted_months = sorted(unique_months.keys())
    last_year, last_month_ord = sorted_months[-1]

    # ✅ include ONLY next month (Option‑2, fixed)
    next_month = last_month_ord + 1
    extra = None
    for i in sheet_items:
        if i["year"] == last_year and MONTH_ORDER.get(i["month"]) == next_month:
            extra = {
                "sheet": i["sheet"],
                "text": "",
                "links": [],
                "month_group": i["month_group"],
                "month_link": i["month_link"],
                "month": i["month"],
                "year": i["year"],
            }
            break

    final_items = matches + ([extra] if extra else [])
    return format_results(final_items)


# ---------------- API ----------------

@app.route("/chat", methods=["POST"])
def chat():
    data = request.get_json(silent=True) or {}
    q = data.get("question", "").strip()
    sheet = data.get("sheet", "").strip()

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
