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

# ---------------- Helpers ----------------

def normalize_spaces(text):
    return re.sub(r"\s+", " ", str(text).strip())


def normalize_month_and_year(text):
    if not text:
        return None, None

    text = str(text).lower()
    text = text.replace(",", " ").replace("-", " ")
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

    parts = text.split()
    month = year = None

    for p in parts:
        if p in months:
            month = months[p]
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


# ---------------- Load Excel ----------------

def load_excel():
    global knowledge
    knowledge.clear()

    wb = load_workbook(EXCEL_PATH, data_only=True)

    for sheet in wb.worksheets:
        cur_text = cur_link = cur_month = cur_year = None

        for row in sheet.iter_rows():

            # ✅ Month detection: Text / Hyperlink / Excel Date (A or B)
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

                if not raw:
                    continue

                txt = normalize_spaces(raw)
                if looks_like_month_group(txt):
                    cur_text = txt
                    cur_link = cell.hyperlink.target if cell.hyperlink else None
                    cur_month, cur_year = normalize_month_and_year(txt)
                    break

            texts, links = [], []

            for cell in row:
                if cell.value:
                    texts.append(normalize_spaces(cell.value))
                if cell.hyperlink and cell.hyperlink.target:
                    links.append(cell.hyperlink.target)

            if not texts:
                continue

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


# ---------------- Formatting (OPTION‑2 ENABLED) ----------------

def format_results(items):
    if not items:
        return "No relevant data found."

    output = []
    seen_months = set()

    for item in items:
        month_key = (item.get("month_group"), item.get("month_link"))

        if item.get("month_group") and month_key not in seen_months:
            seen_months.add(month_key)
            output.append(item["month_group"])
            if item.get("month_link"):
                output.append(item["month_link"])
            output.append("-" * 40)

        if item.get("text"):
            output.append(item["text"])

        for l in item.get("links", []):
            if l != item.get("month_link"):
                output.append(l)

        output.append("")

    return "\n".join(output).strip()


# ---------------- Search Logic (Final Option‑2) ----------------

def search_answer(question, selected_sheet):
    q = question.lower().strip()
    q_month, q_year = normalize_month_and_year(q)

    sheet_items = [i for i in knowledge if i["sheet"].lower() == selected_sheet.lower()]
    if not sheet_items:
        return f"No data found for sheet '{selected_sheet}'."

    # ✅ Keyword search
    words = q.split()
    matches = []

    for item in sheet_items:
        score = sum(1 for w in words if len(w) > 2 and w in item["text"].lower())
        if score:
            matches.append(item)

    if not matches:
        return f"No relevant data found in sheet '{selected_sheet}'."

    # ✅ OPTION‑2: also include EMPTY month buckets (for link display)
    months_in_results = {(i["month"], i["year"]) for i in matches if i["month"] and i["year"]}
    extra_months = []

    for item in sheet_items:
        key = (item.get("month"), item.get("year"))
        if key and key not in months_in_results:
            extra_months.append({
                "sheet": item["sheet"],
                "text": "",
                "links": [],
                "month_group": item["month_group"],
                "month_link": item["month_link"],
                "month": item["month"],
                "year": item["year"],
            })
            months_in_results.add(key)

    return format_results(matches + extra_months)


# ---------------- API ----------------

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
