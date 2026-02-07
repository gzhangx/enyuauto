import * as mainOps from './lib/main_ops';
import *  as freed from './lib/util'

async function test() {    
    const args = process.argv.slice(2);
    let opstr: string | undefined;

    if (args.includes('test')) {
        opstr = 'debug_remove_test_';
    }
    if (args.includes('del')) {
        opstr = 'del';
    }

    const linNumberStr = args.find(arg => arg.match(/^\d+$/));
    if (!linNumberStr) {
        console.error('Please provide a line number as an argument.');
        return;
    }
    const lineNumber = parseInt(linNumberStr);


    const log: mainOps.DebugLog = { 
        operations: [], 
        doLog: (msg: string) => {
          log.operations.push(msg);
          console.log(msg);
        }
    };
    const ops = await mainOps.getSheetOps();
    await mainOps.debug_main(ops, lineNumber, log,opstr).catch((err: Error) => {
        console.error('Test failed:', err);        
    });
    console.log('Log:', log.operations);
}


import * as secs from './enyu_secs.json';
async function quickTest() {
    freed.login({ ...secs.auth }).then((loginData) => {
        console.log('Login successful. Session cookies:', loginData);
    });
}
quickTest();
