import { writeFile } from 'fs/promises';
import { readFile } from 'fs/promises'
import xlsx from 'node-xlsx';
import { resolve } from 'path';
import cliProgress from 'cli-progress';
import workerpool from 'workerpool';


function chunkArray(array: any[], chunkSize: number): any[][] {
	const chunkedArray: any[][] = [];
	let index = 0;

	while (index < array.length) {
		chunkedArray.push(array.slice(index, index + chunkSize));
		index += chunkSize;
	}

	return chunkedArray;
}

const init = async () => {
	
	const pool = workerpool.pool(__dirname + '/worker.js', { maxWorkers: 24, workerType: 'thread' });

	const bar1 = new cliProgress.SingleBar({
		hideCursor: true,
		etaBuffer: 100,
		format: ' {bar} | {value}/{total} | ETA: {eta_formatted}',
	}, cliProgress.Presets.shades_classic);


	const [, , filepath] = process.argv;
	const sheet = xlsx.parse(await readFile(resolve(process.cwd(), filepath)));
	const [page0] = sheet;

	const finalTable: string[][] = [];

	let lines = page0.data.slice(1);
	lines = lines.slice(0, lines.findIndex((cells) => cells.length == 0) + 1);
	
	const allerrors = []
	
	const chunks = chunkArray(lines, 50);
	bar1.start(chunks.length, 0);

	for (const [index, chunk] of chunks.entries()) {
		bar1.update(index)
		const { batchErrors, batchTable } = await pool.exec('loadFiles', [chunk])

		if (batchTable.length > 0) {
			finalTable.push(...batchTable)
		}

		if (batchErrors.length > 0) {
			allerrors.push(...batchErrors)
		}

		await writeFile('errors.xls', xlsx.build([
			{
				name: 'sheet0',
				data: [
					['service site #', 'name of location', 'ar#', 'file path', 'date last invoiced', 'error'],
					...allerrors,
				],
				options: {}
			}
		], {}));

		await writeFile('out.xls', xlsx.build([
			{
				name: 'sheet0',
				data: [
					['service site #', 'name of location', 'ar#', 'file path', 'date last invoiced', 'building name', 'frequency', 'order', 'MFG.', 'SIZE', 'TYPE', 'MFG DATE', 'SERIAL #', 'Last Hydro', 'Last 6yr Test', 'LOCATION', 'STATUS'],
					...finalTable,
				],
				options: {}
			}
		], {}));
	}
	await pool.terminate();
	bar1.stop();
	console.log('DONE')
}

init();