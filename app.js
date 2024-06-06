function generateClaim() {
    const name = document.getElementById('name').value;
    const address = document.getElementById('address').value;
    const phone = document.getElementById('phone').value;
    const email = document.getElementById('email').value;
    const liabelee = document.getElementById('liabelee').value;
    const liabeleeAddress = document.getElementById('liabeleeAddress').value;
    const complaint = document.getElementById('complaint').value;

    const content = `
        <w:p>
            <w:r>
                <w:t>Имя: ${name}</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>Адрес: ${address}</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>Телефон: ${phone}</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>Email: ${email}</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>Претензия:</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>${complaint}</w:t>
            </w:r>
        </w:p>
    `;

    const zip = new JSZip();
    const doc = new window.Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
    });

    const template = `
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            ${content}
        </w:document>
    `;

    zip.file("word/document.xml", template);
    zip.generateAsync({ type: "blob" }).then(function (content) {
        saveAs(content, "Претензия.docx");
    });
}
