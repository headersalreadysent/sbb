const chalk = require('chalk');
const { default: collect } = require('collect.js');
const fs = require('fs');
const Excel = require('exceljs');

class GenerateList {

	command = null;

	types = ['program', 'kurumsal', 'finansman', 'ekonomik', 'hepsi']

	constructor(program) {
		this.command = program.command('doldur')
		this.define();
	}

	define() {
		this.command
			.description('Seçilen tip için şablonları doldurur.')
			.argument('<type>', 'tertip program|kurumsal|finansal|ekonomik')
			.action((type, options) => {
				type = type.toLowerCase();
				if (this.types.indexOf(type) == -1) {
					console.log(chalk.red.bold("Tip seçeneği şunlardan biri olmalıdır: " + this.types.join(', ')))
					return
				}
				if (type == 'hepsi') {
					this.generateAll()
				} else {
					let map = this.generateMap(type)
				}



			});
	}

	generateAll(data) {
		['program', 'kurumsal', 'finansman', 'ekonomik'].forEach((typeItem) => {
			let map = this.generateMap(data, typeItem)
			fs.writeFileSync('./tertip-' + typeItem + '.json', JSON.stringify(map, null, 4))
		})
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
				break
			case 'ekonomik':
				break
		}
	}

	async fillKurumsal() {
		let path = fs.realpathSync('./tertip-kurumsal.json')
		if (!fs.existsSync(path)) {
			return console.log(color.red.bold("Program bütçe kurumsal tertip dosyası oluşturulmadı."))
		}
		let data = JSON.parse(fs.readFileSync(path).toString());
		var fillWith = (item, max, fill) => { return ((fill || "0").repeat(50) + item).slice(max * -1) }
		let rows = [
			["0012", 1, 2022, "Hazine ve Maliye Bakanlığı", "Hazine ve Maliye Bakanlığı",
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
				code, 1, 2022, name, name,
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
			return console.log(color.red.bold("Program bütçe tertip dosyası oluşturulmadı."))
		}
		let data = JSON.parse(fs.readFileSync(path).toString());
		var fillWith = (item, max, fill) => { return ((fill || "0").repeat(50) + item).slice(max * -1) }
		let rows = [];
		let addedCodes = []
		let counts = { program: 0, subprogram: 0, faaliyet: 0, subfaaliyet: 0 }
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
					codeParts[0], "8", 2022, row[0], row[0],
					row[0], row[0], //english
					row[0], row[0] //other
				]);
				addedCodes.push(codeParts[0])
				counts.program++;
			}
			//for sub programme
			let subProgramCode = codeParts[0] + '.' + codeParts[1]
			if (addedCodes.indexOf(subProgramCode) == -1) {
				rows.push([
					subProgramCode, "8", 2022, row[1], row[0] + ' > ' + row[1],
					row[1], row[1], //english
					row[1], row[1] //other
				]);
				addedCodes.push(subProgramCode)
				counts.subprogram++;
			}
			//for faaliyet
			let faaliyetCode = subProgramCode + '.' + codeParts[2]
			if (addedCodes.indexOf(faaliyetCode) == -1) {
				rows.push([
					faaliyetCode, "8", 2022, row[2], row[0] + ' > ' + row[1] + ' > ' + row[2],
					row[2], row[2], //english
					row[2], row[2] //other
				]);
				addedCodes.push(faaliyetCode)
				counts.faaliyet++;
			}
			//for sub faaliyet
			let subFaaliyetCode = faaliyetCode + '.' + codeParts[3]
			if (addedCodes.indexOf(subFaaliyetCode) == -1) {
				rows.push([
					subFaaliyetCode, "8", 2022, row[2],
					row[0] + ' > ' + row[1] + ' > ' + row[2] + ' > ' + row[3],
					row[3], row[3], //english
					row[3], row[3] //other
				]);
				addedCodes.push(subFaaliyetCode)
				counts.subfaaliyet++;
			}

		}
		rows = collect(rows).sort((a, b) => {
			return a[0].length - b[0].length
		}).toArray()
		console.log(chalk.yellow("Program      " + counts.program + ' Adet'));
		console.log(chalk.yellow("Alt Program  " + counts.subprogram + ' Adet'));
		console.log(chalk.yellow("Faaliyet     " + counts.faaliyet + ' Adet'));
		console.log(chalk.yellow("Alt Faaliyet " + counts.subfaaliyet + ' Adet'));
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

	async openExcell(filename) {
		const workbook = new Excel.Workbook();
		await workbook.xlsx.readFile(filename);
		return workbook;
	}
}

module.exports = (program) => {
	return new GenerateList(program)
}