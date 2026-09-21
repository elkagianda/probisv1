fetch("probis.xlsx")
    .then(response => response.arrayBuffer())
    .then(data => {

        const workbook = XLSX.read(data);

        const worksheet =
            workbook.Sheets[workbook.SheetNames[0]];

        const jsonData =
            XLSX.utils.sheet_to_json(worksheet);

        console.log(jsonData);

    });
