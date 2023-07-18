const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');
const axios = require('axios');

// работает только на 100 стр примерно 5 тысяч единиц товара,

// ввод ссылки с выбранными фильтрами !поиск по всем регионам!
// let start_url='https://www.avito.ru/all/odezhda_obuv_aksessuary/obuv_zhenskaya-ASgBAgICAUTeAryp1gI?cd=1&f=ASgBAQECAUTeAryp1gIBQOK8DSTu0TS6q9YCAUXGmgwUeyJmcm9tIjowLCJ0byI6MTAwMH0&q=%D0%BA%D1%80%D0%BE%D1%81%D1%81%D0%BE%D0%B2%D0%BA%D0%B8+%D0%BF%D0%BB%D0%B0%D1%82%D1%84%D0%BE%D1%80%D0%BC%D0%B0';
// let name_book = 'кросы-платформа.xlsx';

// let start_url='https://www.avito.ru/all/odezhda_obuv_aksessuary/obuv_zhenskaya-ASgBAgICAUTeAryp1gI?f=ASgBAQECAUTeAryp1gICQOK8DSTu0TS6q9YCpI0OJMCKlgG8ipYBAUXGmgwUeyJmcm9tIjowLCJ0byI6MTAwMH0&q=%D0%BB%D0%BE%D1%84%D0%B5%D1%80%D1%8B+%D0%BD%D0%B0%D1%82%D1%83%D1%80%D0%B0%D0%BB%D1%8C%D0%BD%D0%B0%D1%8F+%D0%BA%D0%BE%D0%B6%D0%B0';
// let name_book = 'лоферы-натур-кожа.xlsx';

// let start_url='https://www.avito.ru/all/odezhda_obuv_aksessuary/obuv_zhenskaya-ASgBAgICAUTeAryp1gI?cd=1&f=ASgBAQECAUTeAryp1gIBQKSNDjS4ipYBwIqWAbyKlgEBRcaaDBR7ImZyb20iOjAsInRvIjoxMDAwfQ&q=%D1%82%D1%83%D1%84%D0%BB%D0%B8+%D0%BC%D0%B5%D1%80%D0%B8+%D0%B4%D0%B6%D0%B5%D0%B9%D0%BD';
// let name_book = 'туфли+мери+джейн+38+39+40.xlsx';

let start_url='https://www.avito.ru/all/odezhda_obuv_aksessuary/obuv_zhenskaya-ASgBAgICAUTeAryp1gI?d=1&f=ASgBAQECAUTeAryp1gIBQKSNDjTEipYBwIqWAbyKlgEBRcaaDBR7ImZyb20iOjAsInRvIjoxNTAwfQ&q=%D1%82%D1%83%D1%84%D0%BB%D0%B8+%D0%BF%D0%BB%D0%B0%D1%82%D1%84%D0%BE%D1%80%D0%BC%D0%B0+%D0%BA%D0%B0%D0%B1%D0%BB%D1%83%D0%BA+%D0%BA%D0%BE%D0%B6%D0%B0';
let name_book = 'туфли+натуркожа+платформа+каблук+39+40.xlsx';

(async ()=> {
    const page = await call_browser(start_url)
    await exsel_book(page)
})()
function getRandomInt(min, max) {
    min = Math.ceil(min);
    max = Math.floor(max);
    return Math.floor(Math.random() * (max - min) + min); // The maximum is exclusive and the minimum is inclusive
}
async function call_browser(url) {
    const browser = await puppeteer.launch({
        headless: false,
        defaultViewport: false,
        waitForInitialPage: true,
        userDataDir: './tmp',
    });
    const page = await browser.newPage();
    await page.setGeolocation({latitude: 61.25, longitude: 73.41667});
    await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/72.0.3626.121 Safari/537.36');
    await page.emulateCPUThrottling(2)
    await page.setDefaultTimeout(90000)
    await page.goto(url)
    //await page.waitForTimeout(getRandomInt(2000, 5000))
    return page
}
async function scrap(page) {
    let info_list = [];
    let title, price, place, imgset, img, url, size, brend, description, delivery, data, count_from_page, lenght_from_page  =''
    // считаю сколько страниц, чтобы перейходить по ним
    //     const count_buttons = await page.$('[data-marker="pagination-button"]')
    //     const count_button = await page.evaluate((el) => (el).childElementCount, count_buttons)
    //     const num_button = await page.evaluate((el) => (el).innerText, count_buttons)
    //     const many_pages = num_button.indexOf('...')
    //     const next_button = num_button.indexOf('След. →')
    //     let lengt_navigation = 0
    //     if (many_pages != -1) {
    //         lengt_navigation = num_button.substring(many_pages + 3, next_button)
    //     } else {
    //         lengt_navigation = count_button - 2
    //     }
    //     console.log(`количество страниц ${lengt_navigation}`)

    // считаю сколько объявлений найдено
    let count_search = await page.$('span[data-marker="page-title/count"]')
    count_search = await page.evaluate((el)=>(el).textContent, count_search)

    console.log(count_search)
    let p=1
    // пока не пролистали все страницы делаем это
   // while (p<=lengt_navigation) {
    do {
        console.log('i am here')
        await page.waitForTimeout(getRandomInt(2000, 4000))
        //прокрутка страницы вниз для сбора инфы по картинкам
        try{
            count_from_page = await page.$('[data-marker="catalog-serp"]')
            lenght_from_page = await page.evaluate((el) => (el).childElementCount, count_from_page)
            console.log(`на ${p} странице ${lenght_from_page} элемент(-a, -ов)`)
        } catch (e) {}
        if (lenght_from_page!=0) {
            //задаю шаг и скорость прокрутки
            for (let i = 0; i < lenght_from_page; i++) {
                await page.waitForTimeout(getRandomInt(100, 1000));
                await page.mouse.wheel({deltaY: 232});
                await page.focus(".iva-item-content-rejJg")
            }
            //await page.focus('[data-marker="pagination-button"]')

            // список объявлений на текущей странице
            const list_products = await page.$$(".iva-item-content-rejJg")

            // для каждого блока с картинкой и описанием делаем это
            for (let res of list_products) {
                // console.log("считываю информацию")
                try {
                    await page.waitForSelector((".photo-slider-image-YqMGj"))
                    imgset = await page.evaluate((el) => (el).querySelector(".photo-slider-image-YqMGj").getAttribute("srcset"), res)
                    let start_imgset = imgset.indexOf("208w,")
                    let end_imgset = imgset.indexOf('236w,')
                    img = imgset.substring(start_imgset, end_imgset).replace('208w,', '')
                } catch (e) {
                    img = 'https://img.freepik.com/premium-vector/default-image-icon-vector-missing-picture-page-for-website-design-or-mobile-app-no-photo-available_87543-11093.jpg';
                    // img.height = "156";
                    console.log('no img')
                }
                try {
                    url = "https://www.avito.ru" + await page.evaluate((el) => (el).querySelector('.iva-item-sliderLink-uLz1v').getAttribute("href"), res)
                } catch (e) {
                    console.log('no url')
                }
                try{
                    delivery = await page.evaluate((el)=>(el).querySelector(".delivery-icon-gA_bl").tagname, res)
                    delivery = "возможна доставка"
                } catch (e) {delivery='нет'
                }

                title = await page.evaluate((el) => (el).querySelector("div.iva-item-titleStep-pdebR > a > h3").textContent, res)
                let price1 = await page.evaluate((el) => (el).querySelector('.price-price-JP7qe').textContent, res)
                let index_p = price1.indexOf("₽")
                if (index_p!=-1) {
                    let price2 = price1.slice(0, index_p)
                    price2 = price2.replace(/\s/g, '')
                    price = parseInt(price2, 10)
                }
                place = await page.evaluate((el) => (el).querySelector('.geo-root-zPwRk').textContent, res)
                try {
                    let brend_size = await page.evaluate((el) => (el).querySelector('.iva-item-text-Ge6dR.text-text-LurtD.text-size-s-BxGpL').textContent, res)
                    let lenght = brend_size.length
                    let index_sep = brend_size.indexOf(',')
                    if (index_sep != -1) {
                        brend = brend_size.substring(0, index_sep)
                        index_sep = index_sep + 2
                        size = brend_size.substring(index_sep, lenght)
                    } else {
                        brend = ''
                        size = brend_size
                    }
                } catch (e) {
                    console.log('no brend_size')
                }
                try {
                    description = await page.evaluate((el) => (el).querySelector('.iva-item-descriptionStep-C0ty1').textContent, res)
                } catch (e) {
                    description = ''
                    console.log('no description')
                }
                try {
                    data = await page.evaluate((el) => (el).querySelector('.text-text-LurtD.text-size-s-BxGpL.text-color-noaccent-P1Rfs').textContent, res)
                } catch (e) {
                    console.log('no time')
                }

                let info_dict = {
                    img: img,
                    title: title,
                    brend: brend,
                    size: size,
                    price: price,
                    description: description,
                    place: place,
                    delivery: delivery,          
                    data: data,
                    url: url,
                }
                info_list.push(info_dict)

            }
           // await page.focus('[data-marker="pagination-button"]')
        } else {console.log("lenght_from_page==0")}

        console.log(`найдено товаров: ${info_list.length} из ${count_search}`)

        // после того как собрали все данные
        //     if (p<lengt_navigation) {
        //     p = p + 1
        //     let selector = `[data-marker="page(${p})"]`;
        //     try {
        //         await page.click(selector)
        //         console.log(`переход на страницу ${p}`)
        //         await page.waitForNavigation()
        //     }catch (e) {
        //         const body = await page.$('.index-inner-dqBR5.index-innerCatalog-ujLwf')
        //         await body.click({button:'right'})
        //         await page.waitForSelector('.context-menu')
        //         const reload=await page.$('.context-menu li:nth-child(3)')
        //         await reload.click
        //         await page.waitForNavigation()
        //     }
        // } else {break}
        p++
        try {
            await page.click('[aria-label="Следующая страница"]')
            console.log('go')
        } catch (e) {console.log(e)
        }
    } while (!(await page.$('[aria-disabled="true"]')))
    if (info_list.length==count_search){
        console.log("Great!")
    } else {console.log("Oops!")}

    return info_list
}
async function exsel_book(page){
    const info_list=await scrap(page)
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Angel';
    workbook.calcProperties.fullCalcOnLoad = true;
    const sheet=workbook.addWorksheet('MySheet')

    console.log('книга создана')
    console.log('')
    console.log('П О Д О Ж Д И')
    console.log('')

    const headers = Object.keys(info_list[0])
    sheet.addRow(headers)

    headers.forEach((header, index) => {
        const column = sheet.getColumn(index + 1);
    });

    info_list.forEach(item=>{
        const values = headers.map(header=>item[header]);
        sheet.addRow(values);
    })

    console.log('данные загружены')

    sheet.getColumn('A').width = 220 / 7.5;
    sheet.getColumn('B').width = 20
    sheet.getColumn("E").width=9
    sheet.getColumn("G").width=16
    sheet.getColumn("H").width=10
    sheet.getColumn("C").width=10
    sheet.getColumn("D").width=6
    sheet.getColumn("F").width=60
    sheet.getColumn("I").width=10
    sheet.getColumn("J").width=15

    console.log('отображение картинки и привязка ссылки')
   // отображение картинки как картинки, и при нажатии открытие ссылки
   let i =0; let k=1;
    for (let l of info_list){
        let imageUrl = l['img'];
        let imageResponse = await axios.get(imageUrl, { responseType: 'arraybuffer' });
        let imageBuffer = Buffer.from(imageResponse.data, 'binary');

        let imageId = workbook.addImage({
            buffer: imageBuffer,
            extension: 'png',
        });

        sheet.addImage(imageId,
            {tl: { col: 0, row: i + 1 },
                ext: { width: 208, height: 156 },
                hyperlinks: {
                    hyperlink: l['url'],
                    tooltip: l['url']
                }});
        sheet.getRow(k+1).height = 155;
        i++; k++
    }

    console.log('выравнивание данных')
    // выравнение по вертикали по центру, перенос данных
    let column_list = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O' ]
    for (let cl of column_list){
        sheet.eachRow({ includeEmpty: true }, function (row, rowNumber) {
            const cell = row.getCell(cl);
            cell.alignment = { vertical: 'middle', wrapText: true};
        });
    }
    sheet.getRow(1).height = 30;
    sheet.getRow(1).font = {bold:true, color: {argb:'ffffff'}};
    sheet.getRow(1).fill = {type:'pattern',pattern:'solid', fgColor:{argb:"2c974b"}};
    sheet.getRow(1).alignment = { horizontal: 'center' };
    sheet.getRow(1).getCell(1).text.toUpperCase()

       workbook.xlsx.writeFile(name_book)
        .then(()=> {
            console.log('Книга была сохранена!');
        })

}