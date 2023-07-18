const puppeteer = require("puppeteer");
const ExcelJS = require("exceljs");
const axios = require("axios");
const url = require("url");

let start_url = 'https://www.avito.ru/moskva/kvartiry/prodam-ASgBAgICAUSSA8YQ?f=ASgBAQECA0SSA8YQwMENuv036sEN_s45A0DKCDSGWYRZglmO3g4UApDeDhQCA0WECRV7ImZyb20iOjU4LCJ0byI6bnVsbH2sKhR7ImZyb20iOjgsInRvIjpudWxsfcaaDBh7ImZyb20iOjAsInRvIjoxMTAwMDAwMH0';

let name_book = 'квартиры-Москва-11-04-2023.xlsx';

(async ()=> {
    await exsel_book(start_url)
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
    await page.waitForTimeout(getRandomInt(1000, 2000))
    let cookie = await page.cookies()

    let title, price, priceformetr, img, seller_info, seller, count_rooms, sumS, kitchenS,livingS, floor, floor_all, bathroom, otdelka, address, description = ' ';
    let streetAddress, type_rooms, balcony, height, remont, sell, name_building, time_open, type_home, pass_lift, gruz_lift, in_home, yard, parking, year = ' ';
    let metroStations, metroTimes, time_central = '';
    let count_from_page, lenght_from_page = '';
    let info_list =[];

    // считаю сколько страниц, чтобы перейходить по ним
    // const count_buttons = await page.$('[data-marker="pagination-button"]')
    // const count_button = await page.evaluate((el) => (el).childElementCount, count_buttons)
    // const num_button = await page.evaluate((el) => (el).innerText, count_buttons)
    // const many_pages = num_button.indexOf('...')
    // const next_button = num_button.indexOf('След. →')
    // let lengt_navigation = 0
    // if (many_pages != -1) {
    //     lengt_navigation = num_button.substring(many_pages + 3, next_button)
    // } else {
    //     lengt_navigation = count_button - 2
    // }
    // console.log(`количество страниц ${lengt_navigation}`)
    // считаю сколько объявлений найдено
    let count_search = await page.$('span[data-marker="page-title/count"]')
    count_search = await page.evaluate((el)=>(el).textContent, count_search)

    let p=1
    // пока не пролистали все страницы делаем это
 //   while (p<=lengt_navigation) {
    do {
        try {
            count_from_page = await page.$('[data-marker="catalog-serp"]')
            lenght_from_page = await page.evaluate((el) => (el).childElementCount, count_from_page)
            console.log(`на ${p} странице ${lenght_from_page} элемент(-a, -ов)`)
        } catch (e) {
        }

        // считаю сколько объявлений найдено
        let count_search = await page.$('span[data-marker="page-title/count"]')
        count_search = await page.evaluate((el) => (el).textContent, count_search)

        // scrap url from page from results of search
        await page.waitForTimeout(getRandomInt(2000, 4000))
        let url_list = []
        const list_products = await page.$$(".iva-item-content-rejJg")
        for (let res of list_products) {
            const url = "https://www.avito.ru" + await page.evaluate((el) => (el).querySelector('.iva-item-sliderLink-uLz1v').getAttribute("href"), res)
            url_list.push(url)
        }

        //scrap description of url
        for (let url of url_list) {
            let page = await browser.newPage();
            await page.goto(url);
            await page.waitForTimeout(getRandomInt(3000, 9000))
            try {
                let address_metro = await page.$('[itemprop="address"]')
                let address1 = await page.evaluate((el) => (el).outerText, address_metro)
                address = address1.split('\n')

                // ищу адрес, метро и время
                if (address[0]) {
                    streetAddress = address[0];
                }

                const metroRegex = /(\D+)(\d{0,2}[–-]\d{0,2}|от\s?\d{0,2}|до\s?\d{0,2})/;

                for (let i = 1; i < 2; i++) {
                    const match = metroRegex.exec(address[i]);

                    if (match) {
                        metroStations = match[1].trim();

                        const timeRegex = /\d{0,2}[–-]\d{0,2}|от\s?\d{0,2}|до\s?\d{0,2}/;
                        const timeMatch = timeRegex.exec(match[2]);

                        if (timeMatch) {
                            metroTimes = timeMatch[0];
                        }
                    }
                }
                //console.log(metroStations)
                // проверка сколько времени от центpа через Яндекс метро
                if (metroStations!=0) {
                    if(metroStations.length>3) {
                        const newpage = await browser.newPage()
                        await newpage.goto('https://yandex.ru/metro/moscow?scheme_id=sc34974011')
                        await newpage.waitForTimeout(getRandomInt(1000, 1000))
                        await newpage.type('[placeholder="Откуда"]', metroStations, {delay: 50});
                        await newpage.keyboard.press('ArrowDown');
                        await newpage.keyboard.press('Enter');
                        //await newpage.waitForTimeout(getRandomInt(1000, 1000));
                        await newpage.type('[placeholder="Куда"]', 'Площадь Революции', {delay: 50});
                        await newpage.keyboard.press('Enter');
                        //await newpage.waitForTimeout(getRandomInt(1000, 1000));

                        time_central = await newpage.$eval('.masstransit-route-snippet-view__route-duration', (el) => (el).textContent)
                        time_central = parseInt(time_central, 10)

                        await newpage.close()
                    }
                }
            } catch (e) {
                console.log(e)
            } // адрес метро
            try {
                title = await page.$eval(".title-info-title-text", (el) => (el).textContent)
            } catch (e) {
                console.log(e)
            } // title
            try {
                price = await page.$eval('[itemprop="price"]', (el) => (el).textContent)
                price = price.replace(/\s/g, '')
                price = parseInt(price, 10)
            } catch (e) {
                console.log(e)
            } // price
            try {
                priceformetr = await page.$eval('.style-item-price-sub-price-_5RUD', (el) => (el).textContent)
            } catch (e) {
                console.log(e)
            } // priceformetr
            try {
                img = await page.$eval('[data-marker="image-frame/image-wrapper"] > img', (el) => (el).getAttribute('src'))
            } catch (e) {
                console.log(e)
                img = 'https://commons.wikimedia.org/wiki/File:No_image_available_400_x_600.svg#/media/File:No_image_available_600_x_400.svg';
            } // img
            try {
                seller_info = await page.$eval('[data-marker="seller-info/name"]', (el) => (el).textContent)
            } catch (e) {
                console.log(e)
            } // seller_info
            try {
                seller = await page.$eval('[data-marker="seller-info/label"]', (el) => (el).textContent)
            } catch (e) {
                console.log(e)
            } // seller
            try {
                let about_appartment = await page.$(".params-paramsList-zLpAu")
                let about_app = await page.evaluate((el) => (el).outerText, about_appartment)
                about_app = about_app.split('\n')

                let app_dict = {};
                about_app.forEach(item => {
                    const [key, value] = item.split(':');
                    app_dict[key.trim()] = value.trim();
                });

                count_rooms = app_dict['Количество комнат'];
                sumS = app_dict['Общая площадь'];
                sumS = parseFloat(sumS)
                kitchenS = app_dict['Площадь кухни'];
                kitchenS = parseFloat(kitchenS)
                livingS = app_dict['Жилая площадь'];
                livingS = parseFloat(livingS);
                let floors = app_dict['Этаж'];
                let index_floor = floors.indexOf(' из ')
                floor = floors.substring(0, index_floor)
                floor_all = floors.substring(index_floor + 4,)
                bathroom = app_dict['Санузел'];
                otdelka = app_dict['Отделка'];
                type_rooms = app_dict['Тип комнат']
                balcony = app_dict['Балкон или лоджия']
                height = app_dict['Высота потолков']
                remont = app_dict['Ремонт']
                sell = app_dict['Способ продажи']

            } catch (e) {
                console.log(e)
            } // о квартире
            // try {
            //     let address_metro = await page.$('[itemprop="address"]')
            //     let address1 = await page.evaluate((el)=>(el).outerText, address_metro)
            //     address = address1.split('\n')
            //
            //     address.forEach((line, index) => {
            //         if (index === 0) {
            //             streetAddress = line;
            //         } else {
            //             const metroRegex = /(\D+)(\d{0,2}[–-]\d{0,2}|от\s?\d{0,2}|до\s?\d{0,2})/;
            //             const match = metroRegex.exec(line);
            //
            //             if (match) {
            //                 metroStations.push(match[1].trim());
            //
            //                 const timeRegex = /\d{0,2}[–-]\d{0,2}|от\s?\d{0,2}|до\s?\d{0,2}/;
            //                 const timeMatch = timeRegex.exec(match[2]);
            //
            //                 if (timeMatch) {
            //                     metroTimes.push(timeMatch[0]);
            //                 } else {
            //                     metroTimes.push('');
            //                 }
            //             }
            //         }
            //     });
            //     const newpage = await browser.newPage()
            //     await newpage.goto('https://yandex.ru/metro/moscow?scheme_id=sc34974011')
            //     await newpage.waitForTimeout(getRandomInt(1000,3000))
            //
            //     for (let m of metroStations){
            //       await newpage.type('[placeholder="Откуда"]', m, {delay:200});
            //       await newpage.keyboard.press('ArrowDown');
            //       await newpage.keyboard.press('Enter');
            //       await newpage.waitForTimeout(getRandomInt(1000, 3000));
            //       await newpage.type('[placeholder="Куда"]', 'Площадь Революции', {delay:200});
            //       await newpage.keyboard.press('Enter');
            //       await newpage.waitForTimeout(getRandomInt(1000, 3000));
            //       let time_centrals = await newpage.$eval('.masstransit-route-snippet-view__route-duration',(el)=>(el).textContent)
            //       time_central.push(time_centrals)
            //       await newpage.focus('[placeholder="Откуда"]')
            //       await newpage.click('.metro-input-form__stop-suggest._type_from > div > .metro-input-form__input > div > span > .input__icons-right > span > div');
            //       await newpage.focus('[placeholder="Куда"]')
            //       await newpage.click('.metro-input-form__stop-suggest._type_to > div > .metro-input-form__input > div > span > .input__icons-right > span > div')
            //   }
            //     await newpage.close()
            // } catch (e) {
            //     console.log(e)
            // } // адрес метро
            try {
                let about_homes = await page.$(".style-item-params-list-vb1_H")
                let about_home = await page.evaluate((el) => (el).outerText, about_homes)
                about_home = about_home.split('\n')

                let home_dict = {};
                about_home.forEach(item => {
                    const [key, value] = item.split(':');
                    home_dict[key.trim()] = value.trim();
                });

                name_building = home_dict['Название новостройки'];
                type_home = home_dict['Тип дома'];
                time_open = home_dict['Срок сдачи'];
                year = home_dict['Год постройки'];
                year = parseInt(year, 10)
                pass_lift = home_dict['Пассажирский лифт'];
                gruz_lift = home_dict['Грузовой лифт'];
                in_home = home_dict['В доме'];
                yard = home_dict['Двор'];
                parking = home_dict['Парковка'];

            } catch (e) {
                console.log(e)
            } // о доме
            try {
                description = await page.$eval('[itemprop="description"]', (el) => (el).textContent)
            } catch (e) {
            } // description

            let info_dict = {
                img: img,
                info: seller_info,
                seller: seller,
                title: title,
                price: price,
                price_m2: priceformetr,
                rooms: count_rooms,
                Sm2: sumS,
                Sm2kitchen: kitchenS,
                Sm2live: livingS,
                floor: floor,
                floor_all: floor_all,
                streetAddress: streetAddress,
                metroStations: metroStations,
                metroTimes: metroTimes,
                time_central: time_central,
                bathroom: bathroom,
                otdelka: otdelka,
                type_rooms: type_rooms,
                balcony: balcony,
                height: height,
                remont: remont,
                sell: sell,
                name_building: name_building,
                type_home: type_home,
                time_open: time_open,
                year: year,
                pass_lift: pass_lift,
                gruz_lift: gruz_lift,
                in_home: in_home,
                yard: yard,
                parking: parking,
                description: description,
                url: url
            }
            await page.close()

            info_list.push(info_dict)

        }

        console.log(`найдено товаров: ${info_list.length} из ${count_search}`)


        // после того как собрали все данные
        // if (p < lengt_navigation) {
        //     p = p + 1
        //     let selector = `[data-marker="page(${p})"]`;
        //     try {
        //         await page.click(selector)
        //         console.log(`переход на страницу ${p}`)
        //         await page.waitForNavigation()
        //     } catch (e) {
        //         // const body = await page.$('.index-inner-dqBR5.index-innerCatalog-ujLwf')
        //         // await body.click({button: 'right'})
        //         // await page.waitForSelector('.context-menu')
        //         // const reload = await page.$('.context-menu li:nth-child(3)')
        //         // await reload.click
        //         // await page.waitForNavigation()
        //     }
        // } else {
        //     break
        // }
        p++
        await page.click('[aria-label="Следующая страница"]')
    } while (!(await page.$('[aria-disabled="true"]')))


    return info_list
}
async function exsel_book(url){
    const info_list=await call_browser(url)
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
        column.width=10;
    });

    info_list.forEach(item=>{
        const values = headers.map(header=>item[header]);
        sheet.addRow(values);
    })

    console.log('данные загружены')

    sheet.getColumn('A').width = 220 / 7.5;

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
    sheet.columns.forEach(function(column) {
        column.eachCell({ includeEmpty: true }, function(cell, rowNumber) {
            cell.alignment = { vertical: 'middle', wrapText: true };
        });
    });

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

