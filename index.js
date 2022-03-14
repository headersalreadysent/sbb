const { Command } = require('commander');
const program = new Command();

program
	.name('SBBParser')
	.description('SBB Bütçe Parser')
	.version('0.8.0');

require('./actions/cleanCsv')(program);
require('./actions/generateList')(program);


program.parse();