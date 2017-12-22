Kapor.js 
========

NodeJS Excel file parser 

> Only supports xlsx for now.

Install
=======
```sh
npm install kapor
```
Use
====
```js
var kapor = require('kapor');
kapor('Spreadsheet.xlsx').then((data) => {
    // data is an array of arrays
});
```
    
If you use async/await, 
```js
data = await kapor('Spreadsheet.xlsx');
```

Test
=====
Run `npm test`

MIT License.