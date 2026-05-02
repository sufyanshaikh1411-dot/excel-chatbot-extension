window.onload = function () {
    const input = document.getElementById("q");
    const chat = document.getElementById("chat");
    const sendBtn = document.getElementById("sendBtn");

    async function sendMessage() {
        const question = input.value.trim();
        if (!question) {
            return;
        }

        chat.innerHTML += "You: " + question + "\n\n";
        input.value = "";

        try {
            const response = await fetch("http://127.0.0.1:5000/chat", {
                method: "POST",
                headers: {
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({ question: question })
            });

            const data = await response.json();
            chat.innerHTML += "Bot: " + (data.answer || "No answer") + "\n\n";
        } catch (error) {
            chat.innerHTML += "Bot: Error connecting to backend\n\n";
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