const csv = require('csv-parser');
const fs = require('fs')
const officegen = require('officegen')

let results = []; //to save data after reading the file
let filteredData = []; // data ready to write on the file 
let excelData = require('./excelParserFilter')

let familyData = excelData.familyResults();

let regExp = /\(([^)]+)\)/;

function filterData(results) {

    for (let i in results) {
        let obj = results[i];
        let date = results[i]['Date'];
        let date1 = '2018-10-14'; // parameters to filter (dates)
        let date2 = '2018-10-15';
        //saving complete objects
        if (date >= date1 && date <= date2) {
            filteredData.push(obj);
        }
    }

    //take only some properties of the object to write to .docx

    // Create an empty Word object:
    let docx = officegen('docx')

    // Officegen calling this function after finishing to generate the docx document:
    docx.on('finalize', function(written) {
        console.log(
            'Finish to create a Microsoft Word document.'
        )
    })

    // Officegen calling this function to report errors:
    docx.on('error', function(err) {
        console.log(err)
    })

    let objectFormat = {};
    let oldTestVar = '';
    let cleanKeys = [];
    let deleteDuplicates = [];

    filteredData.map(elem => {

        if (elem['Observation Details'] === undefined) {
            elem['Location'] += '';
            elem['Location'].trim();
        } else if (elem['Observation Details'].trim() === 'Heard(s).') {
            elem['Location'] = "*";
        } else {
            elem['Location'] += '';
            elem['Location'].trim();
        }

        const allowed = ['Common Name', 'Scientific Name', 'Location', 'Observation Details'];

        const filtered = Object.keys(elem)
            .filter(key => allowed.includes(key))
            .reduce((obj, key) => {
                return {
                    ...obj,
                    [key]: elem[key]
                };
            }, {});

        cleanKeys.push(filtered)
    })

    deleteDuplicates = cleanKeys.reduce((accumulator, curr) => {

        let name = curr['Common Name'],
            found = accumulator.find(elem => elem['Common Name'] === name)

        if (found) found.Location += ';' + curr.Location;
        else accumulator.push(curr);
        return accumulator;
    }, []);


    deleteDuplicates.map(elem => {

        let myLocation = elem['Location'];

        myLocation = elem['Location'].split(';');

        myLocation = myLocation.filter((item, index) => {
            return myLocation.indexOf(item) === index;
        })

        if (myLocation.length === 1 && myLocation[0] === '*') {
            elem['Scientific Name'] += '*';
            myLocation.unshift('');
            elem['Location'] = myLocation[0];
        } else {
            elem['Location'] = `Seen at: ${myLocation.join()}.`;
        }

        let nameMatch = familyData.find(el => el['English name'] === elem['Common Name']);
        let familyText = '';

        if (nameMatch) {
            familyText = nameMatch.family;
            myArrayFamily = regExp.exec(familyText);
            let testFamilyName = myArrayFamily[1];

            if (oldTestVar !== testFamilyName) {
                oldTestVar = testFamilyName;
                realFamilyName = testFamilyName.toUpperCase();
                objectFormat[realFamilyName] = elem;
            }
        }
    })

    for (let [key, value] of Object.entries(objectFormat)) {
        let familyName = key;
        let commonName = value['Common Name'];
        let scientificName = ` ${value['Scientific Name']}`;
        let locationDetails = value['Location'];

        pObj = docx.createP()
        pObj.addText(familyName, { bold: true, color: '188c18', font_face: 'Calibri', font_size: 16 })
        pObj.addLineBreak()
        pObj.addText(commonName, { bold: true, font_face: 'Calibri', font_size: 12 })
        pObj.addText(scientificName, { bold: true, font_face: 'Calibri', font_size: 12 })
        pObj.addLineBreak()
        pObj.addText(locationDetails, { font_face: 'Calibri', font_size: 12 })
    }



    // Let's generate the Word document into a file:

    let out = fs.createWriteStream('example.docx')

    out.on('error', function(err) {
        console.log(err)
    })

    // Async call to generate the output file:
    docx.generate(out)

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
    .catch(error => console.log(error))