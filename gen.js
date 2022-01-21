const binary = (''.padEnd(30, '1')).padEnd(32, '0')
console.log(binary.substring(0, 8));
const sec1 = parseInt(binary.substring(0, 8), 2)
const sec2 = parseInt(binary.substring(8, 16), 2)
const sec3 = parseInt(binary.substring(16, 24), 2)
const sec4 = parseInt(binary.substring(24, 32), 2)
console.log(`${sec1}.${sec2}.${sec3}.${sec4}`)

let binary2 = `${sec1.toString(2)}${sec2.toString(2)}${sec3.toString(2)}${sec4.toString(2)}`;
 console.log(binary2.match(/1/g).length);