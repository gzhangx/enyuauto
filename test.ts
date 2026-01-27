import * as mainOps from './lib/main_ops';


async function test() {
    const args = process.argv.slice(2);
    let opstr: string | undefined;

    if (args.includes('test')) {
        opstr = 'debug_remove_test_';
    } else if (args.includes('del')) {
        opstr = 'del';
    }

    mainOps.debug_main(opstr).catch((err: Error) => {
        console.error('Test failed:', err);
    });
}

test();
