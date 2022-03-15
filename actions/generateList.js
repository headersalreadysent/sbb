const chalk = require('chalk');
const fs = require('fs');

class GenerateList {

	command = null;

	types = ['program', 'kurumsal', 'finansman', 'ekonomik', 'hepsi']

	constructor(program) {
		this.command = program.command('olustur')
		this.define();
	}

	define() {
		this.command
			.description('olusturulmak istenen tertip tipi içinkaynak dosyayı oluşturur')
			.argument('<type>', 'tertip program|kurumsal|finansal|ekonomik')
			.action((type, options) => {
				if (!fs.existsSync('./odenek.json')) {
					console.log(chalk.red.bold('odenek.json dosyası mevcut değil. csv komutu ile üretmelisiniz.'))
					return
				}
				type = type.toLowerCase();
				if (this.types.indexOf(type) == -1) {
					console.log(chalk.red.bold("Tip seçeneği şunlardan biri olmalıdır: " + this.types.join(', ')))
					return
				}
				let path = fs.realpathSync('./odenek.json')
				var data = JSON.parse(fs.readFileSync(path).toString());
				if (type == 'hepsi') {
					this.generateAll(data)
				} else {
					let map = this.generateMap(data, type)
					fs.writeFileSync('./tertip-' + type + '.json', JSON.stringify(map, null, 4))
				}



			});
	}

	generateAll(data) {
		['program', 'kurumsal', 'finansman', 'ekonomik'].forEach((typeItem) => {
			let map = this.generateMap(data, typeItem)
			fs.writeFileSync('./tertip-' + typeItem + '.json', JSON.stringify(map, null, 4))
		})
	}

	generateMap(data, type) {
		let map = {};
		switch (type) {
			case 'program':
				data.forEach((item) => {
					let code = item['ProgramKod'];
					if (map[code] == undefined) {
						map[code] = [
							item["Program Adı"],
							item["Alt Program Adı"],
							item["Faaliyet Adı"],
							item["Alt Faaliyet Adı"],
						];
					}
				});
				break
			case 'kurumsal':
				data.forEach((item) => {
					let code = item['KurumsalKod'];
					if (map[code] == undefined) {
						map[code] = item["Kurum/Birim"];
					}
				});
				break
			case 'finansman':
				data.forEach((item) => {
					let code = item['FinansmanKod'];
					if (map[code] == undefined) {
						map[code] = item['Finansman'];
					}
				});
				break
			case 'ekonomik':
				data.forEach((item) => {
					let code = item['EkonomikKod'];
					if (map[code] == undefined) {
						map[code] = item['Ekonomik'];
					}
				});
				break
		}
		return map;
	}
}

module.exports = (program) => {
	return new GenerateList(program)
}