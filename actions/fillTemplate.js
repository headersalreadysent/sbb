const chalk = require('chalk');
const { default: collect } = require('collect.js');
const fs = require('fs');
const Excel = require('exceljs');

class GenerateList {

	command = null;

	year = new Date().getFullYear();

	types = ['program', 'kurumsal', 'finansman', 'ekonomik', 'hepsi']

	constructor(program) {
		this.command = program.command('doldur')
		this.define();
	}

	define() {
		this.command
			.description('Seçilen tip için şablonları doldurur.')
			.argument('<type>', 'Tertip program|kurumsal|finansal|ekonomik')
			.option('-y, --year <type>', 'Oluşturulacak şablonlar için yıl değeri', '')
			.action((type, options) => {
				if (options.year != '') {
					this.year = options.year
				}
				type = type.toLowerCase();
				if (this.types.indexOf(type) == -1) {
					console.log(chalk.red.bold("Tip seçeneği şunlardan biri olmalıdır: " + this.types.join(', ')))
					return
				}
				if (type == 'hepsi') {
					this.generateAll()
				} else {
					this.generateMap(type)
				}

			});
	}

	async generateAll() {
		await this.fillProgram();
		await this.fillKurumsal();
		await this.fillFinansman()
		await this.fillEkonomik()
	}

	generateMap(type) {
		switch (type) {
			case 'program':
				this.fillProgram();
				break
			case 'kurumsal':
				this.fillKurumsal();
				break
			case 'finansman':
				this.fillFinansman()
				break
			case 'ekonomik':
				this.fillEkonomik()
				break
		}
	}

	async fillKurumsal() {
		let path = fs.realpathSync('./tertip-kurumsal.json')
		if (!fs.existsSync(path)) {
			return console.log(chalk.red.bold("Program bütçe kurumsal tertip dosyası oluşturulmadı."))
		}
		let data = JSON.parse(fs.readFileSync(path).toString());
		var fillWith = (item, max, fill) => { return ((fill || "0").repeat(50) + item).slice(max * -1) }
		let rows = [
			["0012", 1, this.year, "Hazine ve Maliye Bakanlığı", "Hazine ve Maliye Bakanlığı",
				"Hazine ve Maliye Bakanlığı", "Hazine ve Maliye Bakanlığı",
				"Hazine ve Maliye Bakanlığı", "Hazine ve Maliye Bakanlığı"
			]
		];
		for (let code in data) {
			let name = data[code];
			let codeParts = code.split(".");
			codeParts[0] = fillWith(codeParts[0], 5)
			codeParts[1] = fillWith(codeParts[1], 6)
			rows.push([
				code, 1, this.year, name, name,
				name, name, //english
				name, name //other
			]);

		}
		console.log(chalk.blue.bold(rows.length + " adet kurum kaydı oluşturuldu"))

		let template = fs.realpathSync('./templates/kurumsal.xlsx');
		let book = await this.openExcell(template)
		let worksheet = book.worksheets[0];
		rows.forEach((item, index) => {
			let row = worksheet.getRow(index + 2)
			row.values = item
		});

		let filename = template.replace('/templates/', '/')
		await book.xlsx.writeFile(filename);
		console.log(chalk.green.bold("Kurumsal dosyası oluşturuldu."))
		console.log(chalk.green.bold(filename + " dosyası sisteme yükelenebilir."))

	}

	async fillProgram() {
		let path = fs.realpathSync('./tertip-program.json')
		if (!fs.existsSync(path)) {
			return console.log(chalk.red.bold("Program bütçe tertip dosyası oluşturulmadı."))
		}
		let data = JSON.parse(fs.readFileSync(path).toString());
		var fillWith = (item, max, fill) => { return ((fill || "0").repeat(50) + item).slice(max * -1) }
		let rows = [];
		let addedCodes = []
		for (let code in data) {
			let row = data[code];
			let codeParts = code.split(".");
			codeParts[0] = fillWith(codeParts[0], 5)
			codeParts[1] = fillWith(codeParts[1], 6)
			codeParts[2] = fillWith(codeParts[2], 8)
			codeParts[3] = fillWith(codeParts[3], 8)

			//for program
			if (addedCodes.indexOf(codeParts[0]) == -1) {
				rows.push([
					codeParts[0], "8", this.year, row[0], row[0],
					row[0], row[0], //english
					row[0], row[0] //other
				]);
				addedCodes.push(codeParts[0])
			}
			//for sub programme
			let subProgramCode = codeParts[0] + '.' + codeParts[1]
			if (addedCodes.indexOf(subProgramCode) == -1) {
				rows.push([
					subProgramCode, "8", this.year, row[1], row[0] + ' > ' + row[1],
					row[1], row[1], //english
					row[1], row[1] //other
				]);
				addedCodes.push(subProgramCode)
			}
			//for faaliyet
			let faaliyetCode = subProgramCode + '.' + codeParts[2]
			if (addedCodes.indexOf(faaliyetCode) == -1) {
				rows.push([
					faaliyetCode, "8", this.year, row[2], row[0] + ' > ' + row[1] + ' > ' + row[2],
					row[2], row[2], //english
					row[2], row[2] //other
				]);
				addedCodes.push(faaliyetCode)
			}
			//for sub faaliyet
			let subFaaliyetCode = faaliyetCode + '.' + codeParts[3]
			if (addedCodes.indexOf(subFaaliyetCode) == -1) {
				rows.push([
					subFaaliyetCode, "8", this.year, row[2],
					row[0] + ' > ' + row[1] + ' > ' + row[2] + ' > ' + row[3],
					row[3], row[3], //english
					row[3], row[3] //other
				]);
				addedCodes.push(subFaaliyetCode)
			}

		}
		rows = collect(rows).sort((a, b) => {
			return a[0].length - b[0].length
		}).toArray()
		console.log(chalk.blue.bold(rows.length + " adet program kaydı oluşturuldu"))

		let template = fs.realpathSync('./templates/program.xlsx');
		let book = await this.openExcell(template)
		let worksheet = book.worksheets[0];
		rows.forEach((item, index) => {
			let row = worksheet.getRow(index + 2)
			row.values = item
		});

		let filename = template.replace('/templates/', '/')
		await book.xlsx.writeFile(filename);
		console.log(chalk.green.bold("Program dosyası oluşturuldu."))
		console.log(chalk.green.bold(filename + " dosyası sisteme yükelenebilir."))

	}

	async fillFinansman() {
		let path = fs.realpathSync('./tertip-finansman.json')
		if (!fs.existsSync(path)) {
			return console.log(chalk.red.bold("Bütçe finansman tertip dosyası oluşturulmadı."))
		}
		let data = JSON.parse(fs.readFileSync(path).toString());
		let rows = [];
		for (let code in data) {
			let name = data[code];
			rows.push([
				code, 3, this.year, name, name,
				name, name, //english
				name, name //other
			]);

		}
		console.log(chalk.blue.bold(rows.length + " adet finansman kaydı oluşturuldu"))

		let template = fs.realpathSync('./templates/finansman.xlsx');
		let book = await this.openExcell(template)
		let worksheet = book.worksheets[0];
		rows.forEach((item, index) => {
			let row = worksheet.getRow(index + 2)
			row.values = item
		});

		let filename = template.replace('/templates/', '/')
		await book.xlsx.writeFile(filename);
		console.log(chalk.green.bold("Finansman dosyası oluşturuldu."))
		console.log(chalk.green.bold(filename + " dosyası sisteme yükelenebilir."))

	}

	async fillEkonomik() {
		var fullEko4 = JSON.parse(fs.readFileSync('./eko4.json').toString())
		let path = fs.realpathSync('./tertip-ekonomik.json')
		if (!fs.existsSync(path)) {
			return console.log(chalk.red.bold("Bütçe ekonomik tertip dosyası oluşturulmadı."))
		}
		let data = JSON.parse(fs.readFileSync(path).toString());
		let rows = []
		let sorted = {}
		collect(Object.keys(data)).sort((a, b) => {
			return parseInt(a.replace(/\./g, '')) - parseInt(b.replace(/\./g, ''))
		}).toArray().forEach((code) => {
			sorted[code] = data[code];
		})
		let addedCodes = [];
		for (let code in sorted) {
			let name = data[code];
			let parts = code.split('.');

			//for level 1
			let level1 = parts[0];
			if (fullEko4[level1] != undefined) {
				if (addedCodes.indexOf(level1) == -1) {
					rows.push([
						level1, 4, this.year, fullEko4[level1], fullEko4[level1],
						fullEko4[level1], fullEko4[level1], //english
						fullEko4[level1], fullEko4[level1] //other
					]);
					addedCodes.push(level1)
				}
			} else {
				return console.log(chalk.red.bold("Ekonomik tertip " + level1 + " bulunamadı."))
			}
			//for level 2
			let level2 = parts[0] + '.' + parts[1]
			if (fullEko4[level2] != undefined) {
				if (addedCodes.indexOf(level2) == -1) {
					rows.push([
						level2, 4, this.year, fullEko4[level2], fullEko4[level2],
						fullEko4[level2], fullEko4[level2], //english
						fullEko4[level2], fullEko4[level2] //other
					]);
					addedCodes.push(level2)
				}
			} else {
				return console.log(chalk.red.bold("Ekonomik tertip " + level2 + " bulunamadı."))
			}

			//for level 3
			let level3 = level2 + '.' + parts[2]
			if (fullEko4[level3] != undefined) {
				if (addedCodes.indexOf(level3) == -1) {
					rows.push([
						level3, 4, this.year, fullEko4[level3], fullEko4[level3],
						fullEko4[level3], fullEko4[level3], //english
						fullEko4[level3], fullEko4[level3] //other
					]);
					addedCodes.push(level3)
				}
			} else {
				return console.log(chalk.red.bold("Ekonomik tertip " + level3 + " bulunamadı."))
			}
			//for level 4
			rows.push([
				code, 4, this.year, name, name,
				name, name, //english
				name, name //other
			]);

		}
		console.log(chalk.blue.bold(rows.length + " adet ekonomik kaydı oluşturuldu"))

		let template = fs.realpathSync('./templates/ekonomik.xlsx');
		let book = await this.openExcell(template)
		let worksheet = book.worksheets[0];
		rows.forEach((item, index) => {
			let row = worksheet.getRow(index + 2)
			row.values = item
		});

		let filename = template.replace('/templates/', '/')
		await book.xlsx.writeFile(filename);
		console.log(chalk.green.bold("Ekonomik dosyası oluşturuldu."))
		console.log(chalk.green.bold(filename + " dosyası sisteme yükelenebilir."))
	}

	async openExcell(filename) {
		const workbook = new Excel.Workbook();
		await workbook.xlsx.readFile(filename);
		return workbook;
	}
}

module.exports = (program) => {
	return new GenerateList(program)
}