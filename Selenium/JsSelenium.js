const reader = require('xlsx');
const {By, Key, Builder} = require("selenium-webdriver");
require ("chromedriver");


async function test_case(){

    let driver = await new Builder().forBrowser("chrome").build();

    await driver.get("https://google.com"); 
    
    await driver.manage().window().maximize();

    const file = reader.readFile('./data.xlsx');

    const sleep = (delay) => new Promise((resolve) => setTimeout(resolve, delay));
    var days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    var d = new Date();
    var dayName = days[d.getDay()];    

    const source = file.Sheets[dayName];
    const temp = reader.utils.sheet_to_json(file.Sheets[dayName]);

    var longest = [];
    var shortest = [];
    for (let i = 3; i < temp.length+2; i++) {

        await driver.findElement(By.name("q")).sendKeys(await source[`C${i}`].v);

        await sleep(2000);
    
        let abc = await driver.findElements(By.css("#Alh6id .wM6W7d"));
    
        var word = [];
        for(let n of abc){
            word.push(await n.getText());
        }
        var lword  = word.reduce((a, b) => a.length >= b.length ? a : b);
        var sword = word.reduce((a, b) => a.length <= b.length ? a : b);

        longest.push(lword);
        shortest.push(sword);
        await driver.findElement(By.name("q")).clear();
     }

     driver.quit();
     
     for (let j = 3; j < temp.length+2; j++) {
        reader.utils.sheet_add_aoa(file.Sheets[dayName], [[longest[j-3]]], { origin: `D${j}` });
        reader.utils.sheet_add_aoa(file.Sheets[dayName], [[shortest[j-3]]], { origin: `E${j}` });  
     }
     reader.writeFile(file,'./data.xlsx');
     console.log('Program Completed Successfully');
}

test_case();