// Script pour récupérer le texte et envoyer à GPT
Office.onReady(function(info) {
    if (info.host === Office.HostType.Outlook) {
        Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, function(result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                let emailText = result.value;
                fetchGPTSuggestion(emailText);
            }
        });
    }
});

function fetchGPTSuggestion(emailText) {
    let apiKey = "TON_API_KEY"; // Mets ici ta clé OpenAI
    let prompt = "Complète cette phrase de manière professionnelle : " + emailText;

    fetch("https://api.openai.com/v1/completions", {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            "Authorization": "Bearer " + apiKey
        },
        body: JSON.stringify({
            model: "gpt-4",
            prompt: prompt,
            max_tokens: 50
        })
    })
    .then(response => response.json())
    .then(data => {
        let suggestion = data.choices[0].text.trim();
        insertSuggestion(suggestion);
    })
    .catch(error => console.error("Erreur GPT :", error));
}

function insertSuggestion(suggestion) {
    Office.context.mailbox.item.body.setAsync(suggestion, { coercionType: Office.CoercionType.Text });
}