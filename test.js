
const mainOps = require('./lib/main_ops.js');

mainOps.main().catch(err => {
    console.error('Test failed:', err);
});