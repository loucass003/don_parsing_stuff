import { writeFile } from 'fs/promises';
import { readFile } from 'fs/promises'
import xlsx from 'node-xlsx';
import { resolve } from 'path';
import cliProgress from 'cli-progress';
import workerpool from 'workerpool';

const init = async () => {
	
	const pool = workerpool.pool(__dirname + '/worker.js', { maxWorkers: 256 });

	const bar1 = new cliProgress.SingleBar({
		hideCursor: true,
		etaBuffer: 100,
		format: ' {bar} | {id} | {value}/{total} | ETA: {eta_formatted}',
	}, cliProgress.Presets.shades_classic);


	const [, , filepath] = process.argv;
	const sheet = xlsx.parse(await readFile(resolve(process.cwd(), filepath)));
	const [page0] = sheet;

	const finalTable: string[][] = [];

	const lines = page0.data.slice(1);

	bar1.start(lines.findIndex((cells) => cells.length == 0) + 1, 0);


	const allerrors = []
	for (const [index, line] of lines.entries()) {
		const [id, name, ar, path, date] = line;
		bar1.update(index, { id })
		const { status, data } = await pool.exec('loadFile', [[id, name, ar, path, date], path]);
		if (status === 'ok') {
			finalTable.push(...data as string[][]);
		} else {
			allerrors.push(...data)
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
		}

		if (line.length === 0 || !Number.isInteger(line[0]))
			break;
	}
	
	await pool.terminate();
	bar1.stop();

	await new Promise((resolve) => setTimeout(() => resolve(true), 3000));

	writeFile('out.xls', xlsx.build([
		{
			name: 'sheet0',
			data: [
				['service site #', 'name of location', 'ar#', 'file path', 'date last invoiced', 'building name', 'frequency', 'order', 'MFG.', 'SIZE', 'TYPE', 'MFG DATE', 'SERIAL #', 'Last Hydro', 'Last 6yr Test', 'LOCATION', 'STATUS'],
				...finalTable,
			],
			options: {}
		}
	], {}));

	console.log('DONE')
}

init();