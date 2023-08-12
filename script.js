const nameElement = document.getElementById('name');
const idElement = document.getElementById('id');
const domineElement = document.getElementById('domine');
const h1Element = document.querySelector('.heading h1');
const h3Element = document.querySelector('.heading h3');

let sheetData = [];
const urlParams = new URLSearchParams(window.location.search);
const nameToSearch = urlParams.get('name');
// console.log(nameToSearch1);

fetch('data.xlsx')
    .then(response => response.arrayBuffer())
    .then(data => {
        const workbook = XLSX.read(data, { type: "array" });
        const sheetName = workbook.SheetNames[0];
        sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
        searchFromUrl();
    })
    .catch(error => {
        nameElement.innerHTML = "Error loading data from the Excel file.";
    });

function searchFromUrl() {
    // const urlParams = new URLSearchParams(window.location.search);
    // const nameToSearch = urlParams.get('name');
    
    if (nameToSearch) {
        const results = sheetData.filter(row => row.Name.toLowerCase().includes(nameToSearch.toLowerCase()));
        
        if (results.length > 0) {
            displayData(results);
        } else {
            h1Element.innerHTML = "This Certificate Is Not Verified ❌";
            nameElement.innerHTML = "No data found for the given name.";
            h3Element.style.display = "none";
            idElement.style.display = "none";
            domineElement.style.display = "none";
        }
    } else {
        nameElement.innerHTML = "No name parameter found in the URL.";
        nameElement.innerHTML = "No name parameter found in the URL.";
        idElement.style.display = "none";
        domineElement.style.display = "none";
    }
}

function displayData(data) {
    const result = data[0]; // Assuming there's only one result
    nameElement.innerHTML = `<span>Name</span>       ${result.Name}`;
    idElement.innerHTML = `<span>ID</span>       ${result.ID}`;
    domineElement.innerHTML = `<span>Domin</span>       ${result.Domin}`;
}
