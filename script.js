// Adicionar um novo contato
function addContact() {
    const name = document.getElementById("name").value;
    const email = document.getElementById("email").value;
    const phone = document.getElementById("phone").value;
    const vehicle = document.getElementById("vehicle").value;
    const observations = document.getElementById("observations").value;

    if (name && email && phone) {
        const table = document.getElementById("contacts-table").getElementsByTagName('tbody')[0];
        const newRow = table.insertRow();

        newRow.insertCell(0).innerText = name;
        newRow.insertCell(1).innerText = email;
        newRow.insertCell(2).innerText = phone;
        newRow.insertCell(3).innerText = vehicle;
        newRow.insertCell(4).innerText = observations;

        document.getElementById("contact-form").reset();
    } else {
        alert("Por favor, preencha todos os campos.");
    }
}

// Exportar a tabela para PDF
function exportPDF() {
    const { jsPDF } = window.jspdf;

    const pdf = new jsPDF();
    pdf.text("Ficha de Contato Nissan", 10, 10);

    pdf.autoTable({
        html: "#contacts-table",
        startY: 20,
        headStyles: { fillColor: [192, 0, 0] },
        theme: "grid",
    });

    pdf.save("Contatos_Nissan.pdf");
}

// Exportar a tabela para Excel
function exportExcel() {
    const table = document.getElementById("contacts-table");
    const data = Array.from(table.rows).map(row => Array.from(row.cells).map(cell => cell.innerText));
    const worksheet = XLSX.utils.aoa_to_sheet(data);
    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(workbook, worksheet, "Contatos");
    XLSX.writeFile(workbook, "Contatos_Nissan.xlsx");
}

// Exportar a tabela para Word
function exportToWord() {
    const table = document.getElementById("contacts-table");
    const rows = Array.from(table.rows);
    let content = `<h1>Ficha de Contato Nissan</h1><table border="1" style="border-collapse:collapse;">`;

    rows.forEach(row => {
        content += "<tr>";
        Array.from(row.cells).forEach(cell => {
            content += `<td style="padding:8px;">${cell.innerText}</td>`;
        });
        content += "</tr>";
    });

    content += "</table>";

    const blob = new Blob(['\ufeff' + content], {
        type: "application/msword"
    });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "Contatos_Nissan.doc";
    link.click();
}

// Imprimir a página
function printPage() {
    const tableHTML = document.getElementById("contacts-table").outerHTML;
    const newWindow = window.open("");
    newWindow.document.write("<h1>Ficha de Contato Nissan</h1>");
    newWindow.document.write(tableHTML);
    newWindow.document.close();
    newWindow.print();
}

// Limpar todos os contatos
function clearContacts() {
    if (confirm("Tem certeza que deseja limpar todos os contatos?")) {
        document.querySelector("#contacts-table tbody").innerHTML = "";
    }
}

// Eliminar todos os contatos definitivamente
function deleteAllContacts() {
    if (confirm("Tem certeza de que deseja eliminar todos os contatos definitivamente?")) {
        document.querySelector("#contacts-table tbody").innerHTML = "";
        alert("Todos os contatos foram eliminados.");
    }
}

// Abrir o pop-up de criação de reunião
function openSchedulePopup() {
    // Verificar se já existe um pop-up aberto e removê-lo
    const existingPopup = document.getElementById("schedule-popup");
    if (existingPopup) {
        existingPopup.remove();
    }

    // Criar o novo pop-up
    const popup = document.createElement("div");
    popup.id = "schedule-popup";
    popup.innerHTML = `
        <div class="popup-content">
            <h2>Criar Reunião</h2>
            <input type="text" id="meeting-title" placeholder="Título da Reunião" required>
            <label for="meeting-start">Início:</label>
            <input type="datetime-local" id="meeting-start" required>
            <label for="meeting-end">Fim:</label>
            <input type="datetime-local" id="meeting-end" required>
            <textarea id="meeting-description" placeholder="Descrição (Opcional)"></textarea>
            <button onclick="saveMeeting()">Salvar</button>
            <button onclick="closePopup()">Cancelar</button>
        </div>
    `;
    document.body.appendChild(popup);
}

// Salvar o evento como arquivo .ics
function saveMeeting() {
    const title = document.getElementById("meeting-title").value;
    const start = document.getElementById("meeting-start").value;
    const end = document.getElementById("meeting-end").value;
    const description = document.getElementById("meeting-description").value;

    if (!title || !start || !end) {
        alert("Por favor, preencha todos os campos obrigatórios (Título, Início e Fim).");
        return;
    }

    // Formatar datas no padrão iCalendar
    const startDate = new Date(start).toISOString().replace(/[-:]/g, "").split(".")[0];
    const endDate = new Date(end).toISOString().replace(/[-:]/g, "").split(".")[0];

    // Criar conteúdo do arquivo .ics
    const icsContent = `
BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//Nissan Calendar//PT
BEGIN:VEVENT
UID:${Date.now()}@nissan.com
DTSTAMP:${new Date().toISOString().replace(/[-:]/g, "").split(".")[0]}
DTSTART:${startDate}
DTEND:${endDate}
SUMMARY:${title}
DESCRIPTION:${description}
END:VEVENT
END:VCALENDAR
    `.trim();

    // Criar e baixar o arquivo .ics
    const blob = new Blob([icsContent], { type: "text/calendar" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `${title.replace(/\s+/g, "_")}_meeting.ics`;
    link.click();

    alert("Arquivo de reunião criado e baixado com sucesso!");
    closePopup();
}

// Fechar o pop-up
function closePopup() {
    const popup = document.getElementById("schedule-popup");
    if (popup) popup.remove();
}