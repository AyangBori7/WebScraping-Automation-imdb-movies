//node proj.js --excel=TopRatedMovies.csv --url=https://www.imdb.com/chart/top/?ref_=nv_mv_250

// npm init -y
// npm install puppeteer
// npm install minimist
// npm install excel4node
// npm install jsdom


let minimist = require("minimist");
let puppeteer = require("puppeteer");
let excel = require("excel4node");

let args = minimist(process.argv);

run();
async function run(){
    // start the browser
    let browser = await puppeteer.launch({
        defaultViewport: null,
        args: [
            "--start-maximized"
        ],
        headless: false
    });

    // get a tab
    let pages = await browser.pages();
    let page = pages[0];

    // go to url
    await page.goto(args.url);

    handlePage(browser,page);


}


 async function handlePage(browser,page)
{
    await page.waitForSelector("tbody.lister-list tr td.titleColumn a");

     let urlmovies = await page.$$eval("tbody.lister-list tr td.titleColumn a",function(a){
         let urls = [];

        for(let i=0;i<a.length;i++)
        {
            let url = a[i].getAttribute("href");
            urls.push(url);
        }
        return urls;
     })

     let movies=[];

     for(let i=0;i<5;i++)//Top rated movies sorted according to rating
     { 
         await handlemoviedata(browser,page,urlmovies[i],movies);
     }  
   
     createExcelFile(movies);
     await browser.close();
} 

async function handlemoviedata(browser,page,curl,movies)
{
    let mov={
    };

    let npage = await browser.newPage();
    await npage.goto("https://www.imdb.com" + curl);
    

    await npage.waitForSelector("h1");
    mov.title = await npage.$eval("h1",el => el.innerText);

    await npage.waitForSelector("span.AggregateRatingButton__RatingScore-sc-1ll29m0-1.iTLWoV");
    mov.rating = await npage.$eval("span.AggregateRatingButton__RatingScore-sc-1ll29m0-1.iTLWoV",el => el.innerText);

    await npage.waitForSelector("li.ipc-inline-list__item");
    let elements = await npage.$$eval("li.ipc-inline-list__item",function(a){
        let elem = [];

        let a1 = a[0].innerText;
        let a2 = a[2].innerText;

        elem.push(a1);
        elem.push(a2);

        return elem;

    });

    mov.YoR = elements[0];
    mov.Duration = elements[1];

    await npage.waitForSelector("span.GenresAndPlot__TextContainerBreakpointXL-cum89p-2.gCtawA");
    mov.plot = await npage.$eval("span.GenresAndPlot__TextContainerBreakpointXL-cum89p-2.gCtawA",el => el.innerText);

    movies.push(mov);

     npage.close();
  
}

function createExcelFile(movies) {
    let wb = new excel.Workbook();

    
        let sheet = wb.addWorksheet('Top_Rated_Movies');

        let myStyle = wb.createStyle({
            font: {
              bold: true,
              color: 'FF0000',
            },
          });

        sheet.cell(1, 1).string("Movie Name").style(myStyle);
        sheet.cell(1, 2).string("Rating").style(myStyle);
        sheet.cell(1, 3).string("YoR").style(myStyle);
        sheet.cell(1, 4).string("Duration").style(myStyle);
        sheet.cell(1, 5).string("Plot").style(myStyle);
        
    
        for (let j = 0; j < movies.length; j++) {
            sheet.cell(2 + j, 1).string(movies[j].title);
            sheet.cell(2 + j, 2).string(movies[j].rating);
            sheet.cell(2 + j, 3).string(movies[j].YoR);
            sheet.cell(2 + j, 4).string(movies[j].Duration);
            sheet.cell(2 + j, 5).string(movies[j].plot);
        }
    
    wb.write(args.excel);
}