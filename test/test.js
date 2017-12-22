const kapor  = require('../lib/index');
const assert = require('chai').assert;

describe('Xlsx', function(){
    describe('read column', function(){
        it('read column value C5 and this should equals Z5', async() => {
            const data = await kapor('test/resources/teste.xlsx');
            assert.equal(data[4][2], 'z5', 'value read is not valid');
        });
    });
});
