import xlsx from 'node-xlsx';
import { readFile } from 'fs/promises'
import workerpool from 'workerpool';

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
				if (line.length < 8)
					continue;
				if (!Number.isInteger(line[0]))
					continue;

				const clamped = Math.max(8, Math.min(line.length, 9));				
				table.push([...meta, buildingName, freq, ...line.slice(0, clamped)])
			}
		})
		if (table.length === 0) {
			errors.push([...meta, 'Formatted incorrectly'])
			return { status: 'error', data: errors};
		}
		return { status: 'ok', data: table};
	} catch (e: any) {
		errors.push([...meta, `${e?.message || 'unkown error'}`])
		return { status: 'error', data: errors}
	}
}

const loadFiles = async (lines) => {
	const batchTable = [];
	const batchErrors = []

	const results = await Promise.all(lines.map((line) => {
		const [id, name, ar, path, date] = line;
		return loadFile([id, name, ar, path, date], path);
	}));

	for (const { status, data } of results) {
		if (status === 'ok') {
			batchTable.push(...data as string[][]);
		} else {
			batchErrors.push(...data)
		}
	}
	// // for (const [index, line] of lines.entries()) {

	// const [id, name, ar, path, date] = line;
	// const { status, data } = await loadFile([id, name, ar, path, date], path);
	// 	if (status === 'ok') {
	// 		batchTable.push(...data as string[][]);
	// 	} else {
	// 		batchErrors.push(...data)
	// 	}

	// 	if (line.length === 0 || !Number.isInteger(line[0]))
	// 		break;
	// }

	return {
		batchTable,
		batchErrors
	}
}

workerpool.worker({
	loadFiles
});