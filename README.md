Kapor.js 
========

NodeJS Excel file parser 

> Only supports xlsx for now.

Install
=======
    npm install kapor

Use
====
    var kapor = require('kapor');
    kapor('Spreadsheet.xlsx').then((data) => {
        // data is an array of arrays
    });
    
If you use async/await, 

    data = await kapor('Spreadsheet.xlsx');
    
Test
=====
Run `npm test`

MIT License.