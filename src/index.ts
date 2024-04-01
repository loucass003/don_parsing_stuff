import { writeFile } from 'fs/promises';
import { readFile } from 'fs/promises'
import xlsx from 'node-xlsx';
import { resolve } from 'path';
import cliProgress from 'cli-progress';
import { StaticPool } from 'node-worker-threads-pool';


const init = async () => {
	const allerrors = []

	const loadFile = async (meta: string[], path: string) => {
		const errors = []
		try {
			const sheet = xlsx.parse(await readFile(path));
			const table = [];
			sheet.forEach((page, pageIndex) => {
				if (page.data.length === 0) {
					return;
				}
				if (page.data.length < 4) {
					errors.push([...meta, `Invaid Header on page ${pageIndex + 1}, File is formatated incorectly`])
					return;
				}
				const [line0, _, line2, _2, _3, ...others] = page.data;
				const [buildingName] = line0
				const freq = line2.find((cell) => cell?.includes('FREQUENCY') || cell?.includes('FREQ'))?.replace(/FREQ.*\s+(:|-)\s+(\S+)/, '$2').trim() || 'UNKNOWN';
				for (const [index, line] of others.entries()) {
					if (line.length >= 9)
						continue;
					if (!Number.isInteger(line[0]))
						continue;
					table.push([...meta, buildingName, freq, ...line.slice(0, 9)])
				}
			})
			if (table.length === 0) {
				errors.push([...meta, 'Formatted incorrectly'])
				return errors;
			}
			return table;
		} catch (e: any) {
			errors.push([...meta, `${e?.message || 'unkown error'}`])
			return errors
		}
	}

	const staticPool = new StaticPool({
		size: 16,
		task: loadFile
	});

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

	for (const [index, line] of lines.entries()) {
		const [id, name, ar, path, date] = line;
		bar1.update(index, { id })
		const res = await staticPool.exec([id, name, ar, path, date], path);
		if (res.length === 0) {
			finalTable.push(...res as string[][]);
		} else {
			allerrors.push(...res)
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