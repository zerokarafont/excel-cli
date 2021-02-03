const fs = require('fs')
const _ = require('lodash')
const xlsx = require('node-xlsx').default
const isPrivateIP = require('private-ip')
const Alphabet = require('./config')
const ProgressBar = require('progress')
const chalk = require('chalk')
const ora = require('ora')

const spinner = ora()

const progress = (len) => {
      const bar = new ProgressBar(
        `${chalk.bgBlueBright("正在生成Excel:")} [:bar] :percent`,
        {
          complete: "=",
          incomplete: " ",
          width: 20,
          total: len || 100,
        }
      );

      return bar
}

/**
 * 一期: 抽出源IP， 目的IP， 目的端口， 匹配次数 
 * 二期：
 *      进一步筛选目的地址, 例如筛选出的目的地址是内网的
 *      适配不同厂商的访问控制日志格式
 *      用户可能会导出多张日志表格, 要实现多张表格统一分析
 * @param {string} path 本地文件路径
 */
const parseExcel = (path) => {

  spinner.start(`${chalk.bgBlueBright("开始分析中...")}`);

  // 加载进度条
  const bar = progress();

  bar.tick(10)

  const workSheetsFromFile = xlsx.parse(path);
  // 跳过中文表头
  const data = workSheetsFromFile[0].data.slice(1);

  bar.tick(80)

  // 抽取N~Q列 即 [源地址, 源端口, 目的地址, 目的端口]
  const head = ["服务(目的IP：目的port)", "访问频率", "去向(内网or外网)"];
  // const subHead = ['','源IP', '源端口', '目的IP', '目的端口']
  const needRow = [Alphabet["N"], Alphabet["O"], Alphabet["P"], Alphabet["Q"]];
  const extractData = data
    .map((item) => item.filter((_, idx) => needRow.includes(idx)))
    .filter((item) => item.length);

  // 去重 统计
  const DEST_IP = 2;
  const DEST_PORT = 3;
  const res = _.groupBy(
    extractData,
    (source) => `${source[DEST_IP]}:${source[DEST_PORT]}`
  );

  const buildData = Object.keys(res).reduce(
    (data, key) => {
      const destIP = key.split(":")[0];
      return [
        ...data,
        [
          key, // 服务
          res[key].length, // 访问次数
          isPrivateIP(destIP) ? "内网" : "公网", // 公网 or 内网 IP
        ],
      ];
    },
    [head]
  );

  // 按照访问频率由高到低排序
  const sortData = _.orderBy(buildData, (item) => item[1], ["desc"]);

  const buffer = xlsx.build([{ name: "筛选表.xlsx", data: sortData }]);

  try {
    fs.writeFileSync('防火墙日志分析.xlsx', buffer)
    bar.tick(100)
    spinner.succeed(`${chalk.bold.greenBright("成功生成Excel")}\n`);
  }catch(err) {
      spinner.stop(chalk.bold.red('生成Excel文件出错!'))
      throw new Error(err)
  }
}

module.exports = parseExcel;