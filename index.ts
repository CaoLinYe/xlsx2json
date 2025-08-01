import * as XLSX from 'xlsx';
import * as path from 'path';
import * as fs from 'fs';
import { exit } from 'process';

const in_dir = path.join(__dirname, "../数据表");
const out_dir = path.join(__dirname, "../jsons");

function convertValue(v: any, type: string) {
	if (type == "number") {
		return Number(v);
	}
	else if (type == "list") {
		try {
			return JSON.parse(v);
		} catch (e) {
			console.log("error:", v, type, e);
			return v;
		}
	}
	else {
		return "" + v;
	}
}

function parseDataAndSave(data: any, file: string, jsonName: string) {
	let json_data: any = {};
	let keys = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'x', 'Y', 'Z',];
	let index = 4;
	while (data.hasOwnProperty('A' + index)) {
		let item: any = {};
		for (let i = 0; i < keys.length; ++i) {
			let key = keys[i];
			if (!data[key + 1] && !data[key + 2]) { continue; }
			try {
				let name = data[key + 2].v;
				let type = data[key + 3].v;
				if (data[key + index]) {
					let value = data[key + index].v;
					item[name] = convertValue(value, type);
				}
				else {
					item[name] = null;
				}
			} catch(e) {
				console.log(e);
				console.log(`表${file}第${index}行${key}列报错`);
				exit(1);
			}
		}
		if (jsonName == "HeroLevel.xlsx" || jsonName == "Box.xlsx") {
			json_data[`${data['B' + index].v}_${data['C' + index].v}`] = item;
		} else if (jsonName == "RuneLevel.xlsx") {
			json_data[`${data['B' + index].v}_${data['C' + index].v}_${data['D' + index].v}`] = item;
		} else {
			json_data[data['A' + index].v] = item;
		}
		++index;
	}
	fs.writeFileSync(file, JSON.stringify(json_data));
}

fs.readdir(in_dir, (err: any, list: string[]) => {
	if (!err) {
		for (let item of list) {
			if (item[0] == '~' || item.slice(-4) != "xlsx") { continue; } // 临时文件
			let infile = path.join(in_dir, item);
			let outfile = path.join(out_dir, item.replace('xlsx', 'json'));
			console.log(infile, '->', outfile);
			let data = XLSX.readFile(infile);
			parseDataAndSave(data.Sheets[data.SheetNames[0]], outfile, item);
		}
	}
})
