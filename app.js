const csv = require('csv-parser');
const fs = require('fs')
const officegen = require('officegen')

let results = []; //to save data after reading the file
let filteredData = []; // data ready to write on the file 
let excelData = require('./excelParserFilter')

let familyData = excelData.familyResults();

let regExp = /\(([^)]+)\)/;

let regExpAsterik = new RegExp()

function filterData(results) {

    let count = 0;

    for (let i in results) {
        let obj = results[i];
        let date = results[i]['Date'];
        let date1 = '2019-12-28'; // parameters to filter (dates)
        let date2 = '2019-12-28';
        //saving complete objects
        if (date >= date1 && date <= date2) {
            count++;
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
    let matchArray = [];
    let myArrayOfGroups = [];
    let arrayOfPoppedElem = [];
    let arrayOfFinalGroups = [];

    filteredData.map(elem => {

        //To put * in Location 
        if (elem['Observation Details'] === undefined) {
            elem['Location'] += '';
            elem['Location'].trim();
        } else if (elem['Observation Details'].trim() === 'Heard(s).') {
            elem['Location'] = "*";
        } else {
            elem['Location'] += '';
            elem['Location'].trim();
        }

        //clean the objects to keep just some keys values
        const allowed = ['Common Name', 'Scientific Name', 'Location', 'Observation Details'];

        const filtered = Object.keys(elem)
            .filter(key => allowed.includes(key))
            .reduce((obj, key) => {
                return {
                    ...obj,
                    [key]: elem[key]
                };
            }, {});
        //add into an array 
        cleanKeys.push(filtered)
    })

    //delete some duplicate keys
    deleteDuplicates = cleanKeys.reduce((accumulator, curr) => {

        let name = curr['Common Name'],
            found = accumulator.find(elem => elem['Common Name'] === name)

        if (found) found.Location += ';' + curr.Location;
        else accumulator.push(curr);
        return accumulator;
    }, []);



    //delete repeated locations
    deleteDuplicates.map(elem => {

        let myLocation = elem['Location'];

        //converting a string into array for Location
        myLocation = elem['Location'].split(';');

        //
        myLocation = myLocation.filter((item, index) => {
            return myLocation.indexOf(item) === index;
        })

        if (myLocation.length === 1 && myLocation[0] === '*') {
            elem['Scientific Name'] += '*';
            myLocation.unshift('');
            elem['Location'] = myLocation[0];
        } else if (myLocation.length > 1 && myLocation.indexOf('*') > -1) {
            let index = myLocation.indexOf('*');
            if (index > -1) {
                myLocation.splice(index, 1);
            }
            elem['Location'] = `Seen at: ${myLocation.join()}.`;
        } else {
            elem['Location'] = `Seen at: ${myLocation.join()}.`;
        }

        //match identical elements between both databases base on the Enlgish and Common name
        let nameMatch = familyData.find(el => el['English name'] === elem['Common Name']);
        let familyText = '';

        //creating the final array with the family name
        if (nameMatch) {
            familyText = nameMatch.family;

            if (familyText === '') {
                familyText = '(Others)';
            }
            //finding a match between my array of objects and the familyDataBase 
            let myArrayFamily = regExp.exec(familyText);

            if (myArrayFamily !== null) {
                let testFamilyName = myArrayFamily[1];

                let realFamilyName = testFamilyName.toUpperCase();

                if (oldTestVar !== testFamilyName) {
                    oldTestVar = testFamilyName;

                    //adding the family name with uppercase letters
                    objectFormat[realFamilyName] = new Array();
                }
                objectFormat[realFamilyName].push(elem)
            }

        }
    })


    //matching only species with the content of only Peru  but not others countries or locations outside Peru
    familyData.map(item => {
        let RegExp = /^(?!.*(and|to|Ecuador|Brazil|Bolivia|Argentina|Colombia|Paraguay|Venezuela|Chile|Uruguay|California)).*Peru.*$/

        let myMatch = RegExp.exec(item.range)

        let myScientificName = item['scientific name'];

        if (myMatch !== null) {
            matchArray.push(myScientificName)
        }
    })

    for (key in objectFormat) {

        value = objectFormat[key];

        for (let elem = 0; elem < value.length; elem++) {
            let scientificName = value[elem]['Scientific Name']

            let arrayScientificName = scientificName.split(' ');
            let popped = '';

            if (arrayScientificName.length >= 3) {
                popped = arrayScientificName.pop();

                arrayOfPoppedElem.push(popped);

                let myGroupSpecie = arrayScientificName.join(' ');

                myArrayOfGroups.push(myGroupSpecie);
            }
        }
    }

    for (let i = 0; i < myArrayOfGroups.length - 1; i++) {
        if (myArrayOfGroups[i] === myArrayOfGroups[i + 1]) {
            arrayOfFinalGroups.push(myArrayOfGroups[i])
            arrayOfFinalGroups.push(myArrayOfGroups[i] + ' ' + arrayOfPoppedElem[i])
            arrayOfFinalGroups.push(myArrayOfGroups[i + 1] + ' ' + arrayOfPoppedElem[i + 1])
        }
    }

    console.log(arrayOfFinalGroups);

    let numIndex = 0;

    for (key in objectFormat) {
        let familyName = key;
        pObj = docx.createP()
        pObj.addText(familyName, { bold: true, color: '188c18', font_face: 'Calibri', font_size: 16 })
        pObj.addLineBreak()
        value = objectFormat[key];

        for (let elem = 0; elem < value.length; elem++) {

            let commonName = value[elem]['Common Name'];
            let scientificName = value[elem]['Scientific Name'];
            let locationDetails = value[elem]['Location'];

            numIndex++;

            if (matchArray.includes(scientificName)) {
                pObj.addText('E ', { bold: true, color: 'e71837', font_face: 'Calibri', font_size: 12 })
            }

            if (arrayOfFinalGroups.includes(scientificName)) {

                console.log('Hola')

                let convertToArr = scientificName.split(' ');


                if (convertToArr.length === 2) {
                    pObj.addText(numIndex + '. ', { bold: true, font_face: 'Calibri', font_size: 12 })
                    pObj.addText(commonName, { bold: true, font_face: 'Calibri', font_size: 12 })
                    pObj.addText(' (' + scientificName + ')', { bold: true, font_face: 'Calibri', font_size: 12 })
                    pObj.addLineBreak()
                    pObj.addLineBreak()
                } else {
                    pObj.addText('           ' + commonName + ' - ', { bold: true, font_face: 'Calibri', font_size: 12 })
                    pObj.addText(' (' + scientificName + ')', { bold: true, font_face: 'Calibri', font_size: 12 })
                    pObj.addLineBreak()
                    pObj.addText('           ' + locationDetails, { font_face: 'Calibri', font_size: 12 })
                    pObj.addLineBreak()
                    pObj.addLineBreak()
                }
            } else {

                pObj.addText(numIndex + '. ', { bold: true, font_face: 'Calibri', font_size: 12 })

                if (scientificName.charAt(scientificName.length - 1) === '*') {
                    scientificName = ' (' + scientificName.slice(0, scientificName.length - 1) + ')*';
                } else {
                    scientificName = ' (' + scientificName + ')'
                }

                pObj.addText(commonName, { bold: true, font_face: 'Calibri', font_size: 12 })
                pObj.addText(scientificName, { bold: true, font_face: 'Calibri', font_size: 12 })
                pObj.addLineBreak()
                pObj.addText(locationDetails, { font_face: 'Calibri', font_size: 12 })
                pObj.addLineBreak()
                pObj.addLineBreak()
            }
        }
    }

    /*
    for (let [key, value] of Object.entries(objectFormat)) {

        console.log("key: ", key)

        value.map(elem => console.log(elem))

        let familyName = key;
        let commonName = value['Common Name'];
        let scientificName = value['Scientific Name'];
        let locationDetails = value['Location'];

        if (scientificName === 'Laterallus jamaicensis tuerosi') {
            console.log('yes');
        }
        addEndemicMark = matchArray.filter(elem => elem === scientificName);

        //pObj = docx.createP()
        //pObj.addText(familyName, { bold: true, color: '188c18', font_face: 'Calibri', font_size: 16 })
        //pObj.addLineBreak()
        //pObj.addText(commonName, { bold: true, font_face: 'Calibri', font_size: 12 })
        //pObj.addText(scientificName, { bold: true, font_face: 'Calibri', font_size: 12 })
        //pObj.addLineBreak()
        //pObj.addText(locationDetails, { font_face: 'Calibri', font_size: 12 })
    }
    */

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
        fs.createReadStream('MyEBirdDataFake.csv')
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