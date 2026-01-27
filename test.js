
const mainOps = require('./lib/main_ops.js');

const args = process.argv.slice(2);
let opstr;
if (args.includes('test')) {
    opstr = 'debug_remove_test_';
} else if (args.includes('del')) {
    opstr = 'del';
}

mainOps.main(opstr).catch(err => {
    console.error('Test failed:', err);
});