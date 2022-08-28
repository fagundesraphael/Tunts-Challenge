// import fetch from 'node-fetch';
const fetch = (...args) => import('node-fetch').then(({ default: fetch }) => fetch(...args));
const xl = require('excel4node');
const wb = new xl.Workbook();
const ws = wb.addWorksheet('CountriesList');


// Styling the Excel sheet

// font size and font color

var style = wb.createStyle({
    font: {
        color: '#808080',
        size: 12,
        bold: true
    },
});
var styleTop = wb.createStyle({
    font: {
        color: '#4F4F4F',
        size: 16,
        bold: true
    },
    alignment: {
        wrapText: true,
        horizontal: 'center',
    },
});

var styleNumber = wb.createStyle({
    numberFormat: '#,##0.00',
});

// Styling the excel column
ws.column(1).setWidth(15);
ws.column(3).setWidth(15);
ws.cell(1, 1, 1, 4, true)
    .string("Countries List").style(styleTop);
ws.cell(2, 1)
    .string('Name').style(style);
ws.cell(2, 2)
    .string('Capital').style(style);
ws.cell(2, 3)
    .string('Area').style(style)
ws.cell(2, 4)
    .string('Currencies').style(style);



// getting data from API
console.log("Getting data from API");
const url = "https://restcountries.com/v3.1/all";
fetch(url)
    .then((data) => data.json()).then(countriesInfo => {
        // Setting names in alphabetical order
        countriesInfo.sort(function (a, b) {
            if (a.name.common < b.name.common) { return -1; }
            if (a.name.common > b.name.common) { return 1; }
            return 0;
        });

        // Passing data to excel file

        console.log("Creating excel file");
        let row = 3;
        countriesInfo.forEach(countries => {
            if (!countries.area) {
                console.log(countries);
            }
            ws.cell(row, 1).string(countries.name.common);
            ws.cell(row, 2).string(formatCapital(countries.capital));
            ws.cell(row, 3).number(countries.area || 0).style(styleNumber);
            ws.cell(row, 4).string(formatCurrency(countries.currencies));
            row++;
        });
        wb.write('CountriesList.xlsx');
        console.log("Your excel file is ready");
    }).catch(err => {
        console.log(err);
    });
// formating the data from API

function formatCapital(capital) {
    if (capital && capital.length > 0) {
        return capital.join(', ');
    } else {
        return '-';
    }
}

function formatCurrency(currencies) {
    if (currencies && Object.keys(currencies).length > 0) {
        return Object.keys(currencies).join(", ");
    } else {
        return '-';
    }
}
