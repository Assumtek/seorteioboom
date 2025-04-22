document.addEventListener("DOMContentLoaded", function () {
    const fileInput = document.getElementById("fileInput");
    const button = document.querySelector("button");
    const resultSpan = document.querySelector("p span");
    let participants = [];

    fileInput.addEventListener("change", function (event) {
        const file = event.target.files[0];
        if (file) {
            readExcelFile(file);
        }
    });

    button.addEventListener("click", function () {
        if (participants.length > 0) {
            const winner = participants[Math.floor(Math.random() * participants.length)];
            resultSpan.textContent = `${winner.Nome} - ${winner.Email}`;
        } else {
            alert("Por favor, carregue uma planilha antes de sortear!");
        }
    });

    function readExcelFile(file) {
        const reader = new FileReader();
        reader.readAsArrayBuffer(file);
        reader.onload = function (e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: "array" });
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(sheet);

            if (jsonData.length > 0) {
                // Detecta os nomes corretos das colunas independentemente do formato
                const firstRow = jsonData[0];
                const nameKey = Object.keys(firstRow).find(key => key.toLowerCase().includes("first name"));
                const emailKey = Object.keys(firstRow).find(key => key.toLowerCase().includes("email"));

                if (nameKey && emailKey) {
                    participants = jsonData.map(row => ({ Nome: row[nameKey], Email: row[emailKey] })).filter(row => row.Nome && row.Email);
                    alert("Planilha carregada com sucesso! " + participants.length + " participantes encontrados.");
                } else {
                    alert("A planilha não contém colunas de Nome e Email válidas.");
                }
            } else {
                alert("A planilha está vazia ou não contém os dados corretos.");
            }
        };
    }
});
