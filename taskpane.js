async function summarizeEmail() {
  const item = Office.context.mailbox.item;
  item.body.getAsync("text", async (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const emailText = result.value;
      const summary = await callChatGPT(emailText);
      document.getElementById("output").innerText = summary;
    }
  });
}

async function callChatGPT(emailText) {
  const response = await fetch("http://localhost:3001/api/chatgpt", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ prompt: `Summarize this email:\n${emailText}` })
  });

  const data = await response.json();
  return data.output;
}