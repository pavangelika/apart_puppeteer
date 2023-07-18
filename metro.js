const puppeteer = require("puppeteer");
const fs = require('fs');

let url ='https://ru.wikipedia.org/wiki/%D0%A1%D0%BF%D0%B8%D1%81%D0%BE%D0%BA_%D1%81%D1%82%D0%B0%D0%BD%D1%86%D0%B8%D0%B9_%D0%9C%D0%BE%D1%81%D0%BA%D0%BE%D0%B2%D1%81%D0%BA%D0%BE%D0%B3%D0%BE_%D0%BC%D0%B5%D1%82%D1%80%D0%BE%D0%BF%D0%BE%D0%BB%D0%B8%D1%82%D0%B5%D0%BD%D0%B0';

(async() => {
    const browser = await puppeteer.launch({
        headless: true,
        defaultViewport: false,
        waitForInitialPage: true,
        userDataDir: './tmp',
    });
    const page = await browser.newPage();
    await page.setGeolocation({latitude: 61.25, longitude: 73.41667});
    await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/72.0.3626.121 Safari/537.36');
    await page.emulateCPUThrottling(2)
    await page.goto(url)
    await page.waitForTimeout(2000)
    let cookie = await page.cookies()

    let metro = []
    const metro2 = await page.$$('[style="white-space: nowrap;"]>a')
    for (let m of metro2){
        let metro1 = await page.evaluate((el) => (el).getAttribute('title'), m)
        let index = metro1.indexOf('(')
        metro1 = metro1.substring(0,index-1)
        //metro.push(metro1)
        let stantion = {'stantion':metro1}
        metro.push(stantion)
    }

    // Преобразуем объект в формат JSON
    const jsonData = JSON.stringify(metro);

    // Записываем данные в файл
    fs.writeFile('metro.json', jsonData, (err) => {
        if (err) throw err;
        console.log('Данные успешно сохранены в файл metro.json')
    });

    await browser.close()
})();
