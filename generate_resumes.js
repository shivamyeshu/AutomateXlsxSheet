const ExcelJS = require("exceljs");
const Handlebars = require("handlebars");
const fs = require("fs-extra");
const puppeteer = require("puppeteer");

// Function to read Excel file
async function readExcel(filePath) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.getWorksheet(1);

    let students = [];
    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber > 1) {
            students.push({
                name: row.getCell(1).value,
                email: row.getCell(2).value,
                phone: row.getCell(3).value,
                skills: row.getCell(4).value,
                experience: row.getCell(5).value,
                education: row.getCell(6).value
            });
        }
    });

    return students;
}

// Function to generate HTML from Handlebars template
async function generateHTML(data) {
    const templateSource = await fs.readFile("resume_template.hbs", "utf-8");
    const template = Handlebars.compile(templateSource);
    return template(data);
}

// Function to convert HTML to PDF using Puppeteer
async function generatePDF(html, fileName) {
    const browser = await puppeteer.launch();
    const page = await browser.newPage();
    await page.setContent(html);
    await page.pdf({ path: `resumes/${fileName}.pdf`, format: "A4" });
    await browser.close();
}

// Main function to generate resumes
async function generateResumes() {
    const students = await readExcel("students.xlsx");

    // Create output folder if not exists
    await fs.ensureDir("resumes");

    for (const student of students) {
        const html = await generateHTML(student);
        await generatePDF(html, student.name.replace(/\s+/g, "_"));
        console.log(`Generated resume for: ${student.name}`);
    }

    console.log("All resumes generated successfully!");
}

// Run the script
generateResumes();
