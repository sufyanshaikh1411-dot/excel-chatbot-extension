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
    month = None
    year = None

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
        cur_text = None
        cur_link = None
        cur_month = None
        cur_year = None

        for row in sheet.iter_rows():

            # ✅ Month bucket detection (A or B, text, hyperlink, OR Excel date)
            for cell in row[:2]:
                raw = cell.value

                # ✅ Excel date / datetime
                if isinstance(raw, (datetime, date)):
                    cur_month = raw.strftime("%B").lower()
                    cur_year = str(raw.year)
                    cur_text = raw.strftime("%b-%y")
                    cur_link = cell.hyperlink.target if cell.hyperlink else None
                    break

                # ✅ Hyperlink display text
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

            row_text = []
            row_links = []

            for cell in row:
                if cell.value:
                    row_text.append(normalize_spaces(cell.value))
                if cell.hyperlink and cell.hyperlink.target:
                    row_links.append(cell.hyperlink.target)

            if not row_text:
                continue

            knowledge.append({
                "sheet": sheet.title,
                "text": " | ".join(row_text),
                "links": row_links,
                "month_group": cur_text,
                "month_link": cur_link,
                "month": cur_month,
                "year": cur_year,
            })


load_excel()


# ---------------- Formatting ----------------

def unique_items(items):
    seen = set()
    out = []
    for i in items:
        key = i["sheet"] + "||" + (i["month_group"] or "") + "||" + i["text"]
        if key not in seen:
            seen.add(key)
            out.append(i)
    return out


def format_results(items, limit=5):
    if not items:
        return "No relevant data found."

    first = items[0]
    lines = []

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


# ---------------- Search Logic (FINAL) ----------------

def search_answer(question, selected_sheet):
    q = question.lower().strip()
    q_month, q_year = normalize_month_and_year(q)

    sheet_items = [
        i for i in knowledge
        if i["sheet"].lower() == selected_sheet.lower()
    ]

    if not sheet_items:
        return f"No data found for sheet '{selected_sheet}'."

    # ✅ Month + year → STRICT
    if q_month and q_year:
        matches = [
            i for i in sheet_items
            if i.get("month") == q_month and i.get("year") == q_year
        ]
        matches = unique_items(matches)

        if matches:
            return format_results(matches)

        return f"No updates found for {q_month.title()} {q_year} in sheet '{selected_sheet}'."

    # ✅ Month only → Latest year
    if q_month and not q_year:
        month_items = [i for i in sheet_items if i.get("month") == q_month and i.get("year")]
        if month_items:
            latest_year = max(i["year"] for i in month_items)
            return format_results(
                unique_items([i for i in month_items if i["year"] == latest_year])
            )

    # ✅ Keyword search
    words = q.split()
    scored = []

    for item in sheet_items:
        score = sum(1 for w in words if len(w) > 2 and w in item["text"].lower())
        if score:
            scored.append((score, item))

    if scored:
        scored.sort(key=lambda x: x[0], reverse=True)
        return format_results(unique_items([i for _, i in scored]))

    return f"No relevant data found in sheet '{selected_sheet}'."


# ---------------- API ----------------

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
``
