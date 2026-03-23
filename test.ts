import * as mainOps from './web/shared/main_ops';
import *  as freed from './lib/util'
import * as secs from './enyu_secs.json';
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

    const operations: string[] = [];
    const log: mainOps.DebugLog = { 
        doLog: (msg: string) => {
          operations.push(msg);
          console.log(msg);
        }
    };
    const ops = await mainOps.getSheetOps(secs.gsAuth);
    const fops = mainOps.getFreeCampAndUpdateOperations(freed.freedCampOps);
            const mainCfg = await mainOps.getOpsAndMainList(ops, log);
            let sres: string[] = [];
            const operation = mainOps.getOperationFromLineNumber(mainCfg.operationList, lineNumber);
            if (!operation) {
              throw new Error(`No operation found for line number ${lineNumber}`);
            }
            for (const action of mainCfg.groupAndMainProjectMapping.actions) {
              //const rr = await mainOps.deleteItemActionTask(ops, fops, mainCfg, operation, action, log);
              //sres.push(rr || '');
            }
}



async function quickTest() {
    freed.freedCampOps.login({ ...secs.auth }).then((loginData) => {
        console.log('Login successful. Session cookies:', loginData);
    });
}
quickTest();
