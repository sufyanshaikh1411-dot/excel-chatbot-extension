import pandas as pd
from flask import Flask, request, jsonify
from flask_cors import CORS
from pathlib import Path

app = Flask(__name__)
CORS(app)

EXCEL_PATH = Path(__file__).resolve().parent.parent / "methodology.xlsx"

knowledge = []

def load_excel():
    global knowledge
    knowledge = []

    xls = pd.ExcelFile(EXCEL_PATH)

    for sheet in xls.sheet_names:
        df = xls.parse(sheet)

        for _, row in df.iterrows():
            text = " | ".join([str(v) for v in row if pd.notna(v)])
            if text.strip():
                knowledge.append(text.lower())

load_excel()

def search_answer(question):
    q = question.lower()
    results = []

    for row in knowledge:
        score = sum(1 for word in q.split() if word in row)
        if score > 0:
            results.append((score, row))

    results.sort(reverse=True)

    if not results:
        return "No relevant data found in the Excel file."

    top = results[:3]
    answer = "\n\n".join([r[1] for r in top])

    return answer


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