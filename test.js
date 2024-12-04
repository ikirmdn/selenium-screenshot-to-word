const { Builder, By, Key } = require('selenium-webdriver');
const fs = require('fs');
const path = require('path');
const moment = require('moment-timezone');
const { Document, Packer, Paragraph, ImageRun, PageOrientation } = require('docx');
const { Options } = require('selenium-webdriver/chrome');

async function start() {
    // Tambahkan opsi untuk memastikan browser maximize
    let options = new Options().addArguments('--start-maximized');

    // Inisialisasi WebDriver untuk Chrome
    let driver = await new Builder().forBrowser('chrome').setChromeOptions(options).build();

    try {
        // Pastikan browser tetap maximize
        await driver.manage().window().maximize();

        // Navigasi ke Google
        await driver.get('https://www.google.com/');
        
        // Lakukan pencarian
        await driver.findElement(By.name('q')).sendKeys('Universitas Siliwangi', Key.RETURN);
        await driver.sleep(3000); // Tunggu agar halaman dimuat sepenuhnya

        // Ambil tangkapan layar
        let screenshot = await driver.takeScreenshot();

        // Tentukan nama file screenshot dengan timestamp Jakarta
        const timestamp = moment().tz('Asia/Jakarta').format('YYYY-MM-DD_HH-mm-ss');
        const screenshotName = `screenshot-${timestamp}.png`;
        const screenshotPath = path.join(__dirname, screenshotName);

        // Simpan tangkapan layar
        fs.writeFileSync(screenshotPath, screenshot, 'base64');
        console.log(`Screenshot saved at: ${screenshotPath}`);

        // Buat dokumen Word baru dengan orientasi landscape
        const doc = new Document({
            sections: [
                {
                    properties: {
                        page: {
                            size: {
                                orientation: PageOrientation.LANDSCAPE, // Set halaman menjadi landscape
                            },
                        },
                    },
                    children: [
                        new Paragraph("Screenshot dari hasil pencarian Google:"),
                        new Paragraph({
                            children: [
                                new ImageRun({
                                    data: fs.readFileSync(screenshotPath),
                                    transformation: {
                                        width: 698.5, // Lebar gambar dalam pt (24.57 cm)
                                        height: 392.9, // Tinggi gambar dalam pt (13.82 cm)
                                    },
                                }),
                            ],
                        }),
                    ],
                },
            ],
        });

        // Simpan dokumen Word
        const wordFilePath = path.join(__dirname, `GoogleSearchResults-${timestamp}.docx`);
        const buffer = await Packer.toBuffer(doc);
        fs.writeFileSync(wordFilePath, buffer);

        console.log(`Word document saved at: ${wordFilePath}`);

        // Tunggu 10 detik sebelum browser ditutup
        await driver.sleep(10000);
    } catch (err) {
        console.error('Terjadi kesalahan:', err);
    } finally {
        // Tutup browser
        await driver.quit();
    }
}

start();
