const puppeteer = require('puppeteer-extra');
const StealthPlugin = require('puppeteer-extra-plugin-stealth');
const xlsx = require('xlsx');

puppeteer.use(StealthPlugin());

const url = 'https://www.amazon.in/s?bbn=81107433031&rh=n%3A81107433031%2Cp_85%3A10440599031&_encoding=UTF8';

async function fetchData() {
    const browser = await puppeteer.launch({ headless: true });
    const page = await browser.newPage();
    await page.goto(url, { waitUntil: 'networkidle2' });

    const products = await page.evaluate(() => {
        const items = document.querySelectorAll('.a-section.a-spacing-small.puis-padding-left-small.puis-padding-right-small');
        const productArray = [];

        items.forEach(item => {
            const name = item.querySelector('.a-section.a-spacing-none.a-spacing-top-small.s-title-instructions-style h2 a span.a-size-base-plus.a-color-base.a-text-normal')?.innerText.trim();
            const price = item.querySelector('.a-price-whole')?.innerText.trim();
            const dileveryBy = item.querySelector('.a-color-base.a-text-bold')?.innerText.trim();
            const rating = item.querySelector('.a-icon.a-icon-star-small.a-star-small-4.aok-align-bottom span.a-icon-alt')?.innerText.trim() || 'No Rating';

            productArray.push({ name, price, dileveryBy, rating });
        });

        return productArray;
    });

    await browser.close();
    return products;
}

async function saveToExcel(products) {
    const workbook = xlsx.utils.book_new();
    const worksheet = xlsx.utils.json_to_sheet(products);

    xlsx.utils.book_append_sheet(workbook, worksheet, 'Products');
    xlsx.writeFile(workbook, 'products.xlsx');

    console.log('Data saved to products.xlsx');
}

(async () => {
    const products = await fetchData();
    if (products.length > 0) {
        await saveToExcel(products);
    } else {
        console.log('No products found or failed to scrape.');
    }
})();
