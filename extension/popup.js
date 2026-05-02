window.onload = function () {
    const input = document.getElementById("q");
    const chat = document.getElementById("chat");
    const sendBtn = document.getElementById("sendBtn");
    const sheetSelect = document.getElementById("sheetSelect");

    function formatLinks(text) {
        const urlRegex = /(https?:\/\/[^\s]+)/g;
        return text.replace(urlRegex, function (url) {
            return `<a href="${url}" target="_blank">${url}</a>`;
        });
    }

    async function sendMessage() {
        const question = input.value.trim();
        const sheet = sheetSelect.value;

        if (!question) return;

        chat.innerHTML += `<div class="user-msg"><b>You:</b> [${sheet}] ${question}</div>`;
        input.value = "";

        try {
            const response = await fetch("https://excel-chatbot-extension.onrender.com/chat", {
                method: "POST",
                headers: {
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({
                    question: question,
                    sheet: sheet
                })
            });

            const data = await response.json();
            const formattedAnswer = formatLinks(data.answer || "No answer found.");

            chat.innerHTML += `<div class="bot-msg"><b>Bot:</b><br>${formattedAnswer}</div>`;
        } catch (error) {
            chat.innerHTML += `<div class="bot-msg"><b>Bot:</b> Error connecting to backend</div>`;
            console.error(error);
        }

        chat.scrollTop = chat.scrollHeight;
    }

    sendBtn.addEventListener("click", sendMessage);

    input.addEventListener("keydown", function (e) {
        if (e.key === "Enter") {
            sendMessage();
        }
    });
};
