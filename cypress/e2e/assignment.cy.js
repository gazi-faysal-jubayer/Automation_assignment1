const XLSX = require('xlsx');

const currentDate = new Date();

// Define options for formatting the date
const options = { weekday: 'long' };

// Get the day name
const dayName = currentDate.toLocaleDateString('en-US', options);

Cypress.Commands.add('readExcelData', (filePath) => {
    cy.readFile(filePath, 'binary').then((fileContent) => {
        const workbook = XLSX.read(fileContent, { type: 'binary' });
        const worksheet = workbook.Sheets[dayName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        // Use `jsonData` for further processing or assertions

        // var index1 = 3;
        // Perform assertions or further processing with the data
        for (let index = 3; index < 13; index++) {
            const id = worksheet[`C${index}`].v;
            cy.visit('https://www.google.com/')

            // change language to English
            const changeLanguage = cy.get('#SIvCob > a').invoke('text').then((text) => {
                if (text == 'English') {
                    cy.get('#SIvCob > a').click()
                }
            });

            // type on the shearch bar
            cy.get('#APjFqb').type(id);


            // Find the suggestion box using its unique class
            cy.get('.erkvQe').within(() => {
                // Get all the suggestion elements
                cy.get('li').then((suggestions) => {
                    let longestText = '';
                    let shortestText = '';

                    suggestions.each((index, suggestion) => {
                        const text = Cypress.$(suggestion).find('.wM6W7d span').text().trim();

                        // Ignore empty text
                        if (text !== '') {
                            // Check for longest text
                            if (text.length > longestText.length) {
                                longestText = text;
                            }

                            // Check for shortest text
                            if (shortestText === '' || text.length < shortestText.length) {
                                shortestText = text;
                            }
                        }
                        // Use the longest and shortest text as needed
                    });
                    // // Use the longest and shortest text as needed
                    worksheet[`D${index}`] = { t: 's', v: `${longestText}` };
                    worksheet[`E${index}`] = { t: 's', v: `${shortestText}` };
                    // index1=index1+1;
                });
            });
            
        }
        cy.wrap(jsonData).as('excelData');
        cy.wrap(workbook).as('workbook');
    });
});

it('reads Excel data', () => {
    cy.readExcelData('cypress/fixtures/Excel.xlsx');

    // Access the Excel data using the `excelData` alias
    cy.get('@workbook').then((workbook) => {
        const fileContent = XLSX.write(workbook, { bookType: 'xlsx', type: 'binary' });
        const filePath = 'updatedFile.xlsx'; // Specify the desired file path
        cy.writeFile(filePath, fileContent, 'binary');

    });
});