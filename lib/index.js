const xlsx       = require('xlsx');
const vetColumns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];

const createVector = (qty) => {
    return Array.apply(null, Array(qty));
}

const read     = (sheets, plan) => {
    const ref  = sheets[plan]['!ref'].split(':');
    const vCol = 'v';
    let start  = ref[0].split(/(\d+)/).filter(Boolean);
    let end    = ref[1].split(/(\d+)/).filter(Boolean);
    start[1]   = parseInt(start[1]);
    end[1]     = parseInt(end[1]);
    const dif  = vetColumns.indexOf(end[0]) - vetColumns.indexOf(start[0]) + 1;

    return createVector(end[1]).map((item, index) => {
        item = {}; index+=1;
        createVector(dif).map((_, letterIndex) => {
            let letter        = vetColumns[letterIndex];
            item[letterIndex] = sheets[plan][letter.concat(index)][vCol];
        });
        return item;
    });
}

module.exports = (path) => {
    return new Promise((resolve, reject) => {
        let xmlsFile = xlsx.readFile(path, null);
        const plan   = xmlsFile.Props.SheetNames;
        const result = read(xmlsFile.Sheets, plan[0]);
        resolve(result);
    });
}

