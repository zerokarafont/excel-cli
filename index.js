#!/usr/bin/env node

const { program } = require('commander')
const chalk = require('chalk')

program.version('0.0.1')

program
    .command('excel')
    .description('筛选过滤并生成新的Excel')
    .option('-p, --path <path>', '本地Excel文件路径')
    .action((cmd) => {
        if (!cmd.path) {
            console.log(chalk.bold.red("缺少本地文件路径参数！"));
            process.exit(1)
        }
        const excel = require('./excel')

        excel(cmd.path)
    })

// 修改 help 提示信息
program._helpDescription = '输出帮助信息'
program.on('--help', function () {
    const tips = `\nTips:
${chalk.gray("–")} 查看帮助
  ${chalk.cyan("sangfor -h")}
  ${chalk.cyan("sangfor excel -h")}
${chalk.gray("–")} 用法
  ${chalk.cyan("sangfor excel -p")} ${chalk.magenta("<本地文件路径 支持直接将文件拖进命令行窗口>")}
${chalk.gray("–")} 用法事例
  ${chalk.cyan("sangfor excel -p  /home/tmp/测试.xlsx")}`;
    console.log(tips)
})

program.parse(process.argv)
