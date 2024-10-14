const { Builder, By, Key } = require('selenium-webdriver');
const fs = require('fs');
const path = require('path');
const moment = require('moment-timezone');
const { Document, Packer, Paragraph, ImageRun } = require('docx');

async function start() {
    // Inisialisasi WebDriver untuk browser Chrome
    let driver = await new Builder().forBrowser('chrome').build();

    try {
        // Navigasi ke Google
        await driver.get('https://www.google.com/');
        
        // Lakukan pencarian
        await driver.findElement(By.name('q')).sendKeys('Universitas Siliwangi', Key.RETURN);
        await driver.sleep(3000); // Tunggu agar halaman hasil pencarian dimuat sepenuhnya

        // Ambil tangkapan layar (screenshot)
        let screenshot = await driver.takeScreenshot();

        // Dapatkan timestamp sesuai zona waktu Jakarta
        const timestamp = moment().tz('Asia/Jakarta').format('YYYY-MM-DD_HH-mm-ss'); // Formatkan timestamp

        // Tentukan nama file screenshot dengan timestamp Jakarta
        const screenshotName = `screenshot-${timestamp}.png`;

        // Tentukan path untuk menyimpan screenshot
        const screenshotPath = path.join(__dirname, screenshotName);

        // Simpan tangkapan layar sebagai file PNG
        fs.writeFileSync(screenshotPath, screenshot, 'base64');
        console.log(`Screenshot saved at: ${screenshotPath}`);

        // Buat dokumen Word baru
        const doc = new Document({
            sections: [{ // Pastikan ada array sections yang valid
                properties: {},
                children: [
                    new Paragraph("Screenshot dari hasil pencarian Google:"),
                    new Paragraph({
                        children: [
                            new ImageRun({
                                data: fs.readFileSync(screenshotPath),
                                transformation: {
                                    width: 600,  // Atur lebar gambar
                                    height: 550, // Atur tinggi gambar
                                },
                            }),
                        ],
                    }),
                ],
            }],
        });

        // Simpan dokumen Word sebagai file
        const wordFilePath = path.join(__dirname, `GoogleSearchResults-${timestamp}.docx`);
        const buffer = await Packer.toBuffer(doc);
        fs.writeFileSync(wordFilePath, buffer);

        console.log(`Word document saved at: ${wordFilePath}`);
    } catch (err) {
        console.error('Terjadi kesalahan:', err);
    } finally {
        // Tutup browser
        await driver.quit();
    }
}

start();
