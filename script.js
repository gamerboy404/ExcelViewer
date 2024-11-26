document.getElementById('excelFile').addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (!file) return;

    const formData = new FormData();
    formData.append('file', file);

    fetch('http://localhost:5123/api/excel/upload', {
        method: 'POST',
        body: formData,
    })
        .then(response => response.json())
        .then(data => {
            const tableHead = document.getElementById('tableHead');
            const tableBody = document.getElementById('tableBody');
            tableHead.innerHTML = '';
            tableBody.innerHTML = '';

            if (data.length) {
                // Create header row
                Object.keys(data[0]).forEach(col => {
                    const th = document.createElement('th');
                    th.textContent = col;
                    tableHead.appendChild(th);
                });

                // Create data rows
                data.forEach(row => {
                    const tr = document.createElement('tr');
                    Object.values(row).forEach(cell => {
                        const td = document.createElement('td');
                        td.textContent = cell || '';
                        tr.appendChild(td);
                    });
                    tableBody.appendChild(tr);
                });
            }
        })
        .catch(error => {
            console.error('Error:', error);
        });
});
