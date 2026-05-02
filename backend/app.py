from pathlib import Path
from flask import Flask, request, jsonify
from flask_cors import CORS
from openpyxl import load_workbook

app = Flask(__name__)
CORS(app)

EXCEL_PATH = Path(__file__).resolve().parent.parent / "methodology.xlsx"

knowledge = []


def load_excel():
    global knowledge
    knowledge = []

    wb = load_workbook(EXCEL_PATH, data_only=True)

    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            text_parts = []
            links = []

            for cell in row:
                if cell.value:
                    text_parts.append(str(cell.value))

                # 👇 THIS extracts hyperlink
                if cell.hyperlink:
                    links.append(cell.hyperlink.target)

            if text_parts or links:
                knowledge.append({
                    "text": " | ".join(text_parts),
                    "links": links
                })


load_excel()


def search_answer(question):
    q = question.lower()
    results = []

    for item in knowledge:
        searchable = item["text"].lower()
        score = sum(1 for word in q.split() if word in searchable)

        if score > 0:
            results.append((score, item))

    results.sort(key=lambda x: x[0], reverse=True)

    if not results:
        return "No relevant data found."

    top = results[:3]
    output = []

    for score, item in top:
        line = item["text"]

        # 👇 Append links
        if item["links"]:
            for link in item["links"]:
                line += f"\n{link}"

        output.append(line)

    return "\n\n".join(output)


@app.route("/chat", methods=["POST"])
def chat():
    data = request.get_json()
    question = data.get("question", "")

    answer = search_answer(question)
    return jsonify({"answer": answer})


if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
