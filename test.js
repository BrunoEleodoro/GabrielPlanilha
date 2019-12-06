var fs = require('fs');

var contents = fs.readFileSync("a.csv", { encoding: 'utf8' });

console.log(contents.toString().split("\n")[32].toString())


// console.log(contents.toString().split("\n")[941].toString().split(",").length)
// console.log(contents.toString().split("\n")[942].toString().split(",").length)
// console.log(contents.toString().split("\n")[943].toString().split(",").length)

