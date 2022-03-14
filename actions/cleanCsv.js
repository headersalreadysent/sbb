const chalk = require('chalk');
const csv = require('csv-parser')
const fs = require('fs')

class CleanCsv {

	command = null;

	constructor(program) {
		this.command = program.command('csv')
		this.define();
	}

	define() {
		this.command
			.description('Csv dosyasını temizleyerek json kaynağı oluşturur')
			.argument('<path>', 'Csv file path')
			.action((path, options) => {
				if (!fs.existsSync(path)) {
					console.log(chalk.red.bold(path + ' dosyası mevcut değil'))
					return
				}
				path=fs.realpathSync(path)
				var jsonPath=path.split('/');
				jsonPath=jsonPath.slice(0,-1).join('/')+'/odenek.json';
				let year = new Date().getFullYear();
				let numberCols = [year, year + 1, year + 2, (year - 1) + " YSHT", (year) + " Tavan Üstü", (year - 1) + " KBÖ"]
				var results = [];
				fs.createReadStream(path)
					.pipe(csv())
					.on('data', (data) => {
						if (!data[''].startsWith("->")) {
							numberCols.forEach((item) => {
								data[item+' Text']=data[item];
								data[item] = parseFloat(data[item].replace(/,/g, '').replace('.', ','))
							})
							delete data[''];
							delete data['İşlemler']
							let tertipParts = data['Tertip'].split('-');
							data['ProgramKod'] = tertipParts[0]
							data['KurumsalKod'] = tertipParts[1]
							data['FinansmanKod'] = tertipParts[2]
							data['EkonomikKod'] = tertipParts[3]
							results.push(data);
						}
					})
					.on('end', () => {
						fs.writeFileSync(jsonPath,JSON.stringify(results,null,4));
						console.log(chalk.green.bold("Öndenek csv->json dönüşlümü tamamlandı."))
						console.log(chalk.green.bold("odenek.json dosyası oluşturuldu."))
						console.log(chalk.blue(results.length+' satır ödenek'))


					});
			});
	}
}

module.exports = (program) => {
	return new CleanCsv(program)
}