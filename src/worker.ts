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
				if (line.length >= 9)
					continue;
				if (!Number.isInteger(line[0]))
					continue;
				table.push([...meta, buildingName, freq, ...line.slice(0, 9)])
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

workerpool.worker({
	loadFile,
});