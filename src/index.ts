import { writeFile } from 'fs/promises';
import { readFile } from 'fs/promises'
import xlsx from 'node-xlsx';
import { resolve } from 'path';
import cliProgress from 'cli-progress';

const loadFile = async (meta: string[], path: string) => {
	try {
		const sheet = xlsx.parse(await readFile(path));
		const [page0] = sheet;
		const [line0, _, line2, _2, _3, ...others] = page0.data;
		const [buildingName] = line0
		const freq = line2.find((cell) => cell?.includes('FREQUENCY')).replace(/FREQUENCY:\s+(\S+)/, '$1');
		const table = [];
		for (const [index, line] of others.entries()) {
			if (line.length !== 9)
				continue;
			if (!Number.isInteger(line[0]))
				continue;
			table.push([...meta, buildingName, freq, ...line])
		}
		if (table.length === 0)
			return null;
		return table;
	} catch {
		return null
	}
}

const init = async () => {
	const bar1 = new cliProgress.SingleBar({
		hideCursor: true,
		format: ' {bar} | {id} | {value}/{total} | ETA: {eta_formatted}',
	}, cliProgress.Presets.shades_classic);


	const [, , filepath] = process.argv;
	const sheet = xlsx.parse(await readFile(resolve(process.cwd(), filepath)));
	const [page0] = sheet;

	const finalTable: string[][] = [];

	const lines = page0.data.slice(1);

	bar1.start(lines.length, 0);

	for (const [index, line] of lines.entries()) {
		const [id, name, ar, path, date] = line;
		bar1.update(index, { id })
		const res = await loadFile([id, name, ar, path, date], path);
		if (res) {
			finalTable.push(...res as string[][]);
		} else {
			console.log(`Error loading ${path}`);
		}
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
}

init();