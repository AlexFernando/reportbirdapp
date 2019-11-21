const csv = require('csv-parser');
const fs = require('fs');
//package to create .docx
const { Document, Packer, Paragraph, TextRun } = require('docx');
// Create document
const doc = new Document();

let results = []; //to save data after reading the file
let filteredData = []; // data ready to write on the file 


function filterData(results) {

    for (let i in results) {
        let obj = results[i];
        let date = results[i]['Date'];
        let date1 = '2018-04-01'; // parameters to filter (dates)
        let date2 = '2018-04-31';
        //saving complete objects
        if (date >= date1 && date <= date2) {
            filteredData.push(obj);
        }
    }

    //take only some properties of the object to write to .docx

    filteredData.map(elem => {

        let scientificName = `(${elem['Scientific Name']})`;
        let commonName = `${elem['Common Name']} `;
        let locationDetails = '';

        if (elem['Observation Details'] === undefined) {
            locationDetails = `Seen at ${elem.Location}`
        } else if (elem['Observation Details'].trim() === 'Heard(s).') {
            scientificName += "*";
            locationDetails = "";
        } else {
            locationDetails = `Seen at ${elem.Location}`;
        }

        // actually writing to the new .docx
        doc.addSection({
            properties: {},
            children: [
                new Paragraph({
                    children: [
                        new TextRun({
                            text: commonName,
                            bold: true,
                        }),
                        new TextRun({
                            text: scientificName,
                            bold: true
                        }),
                        new TextRun(locationDetails).break(),
                    ],
                }),
            ],
        });

        // Used to export the file into a .docx file
        Packer.toBuffer(doc).then((buffer) => {
            fs.writeFileSync("My Document.docx", buffer);
        });
    })

    return filteredData;
}

// return a Promise
const readFilePromise = () => {
    return new Promise((resolve, reject) => {
        fs.createReadStream('MyEBirdData.csv')
            .pipe(csv())
            .on('data', data => results.push(data))
            .on('end', () => {
                resolve(results);
            });
    })
}

//handling the Promise and using filterData function 
readFilePromise()
    .then(result => filterData(result))
    .catch(error => console.log(error));