const io = require('socket.io-client');
const fs = require('fs');
const XLSX = require('xlsx');

const filePath = 'sheet_Std.xlsx'; // Path to your local spreadsheet
const outputFilePath = 'conversation_std.txt';
const userId = '6683ae3e0f39a8f05bfcef03'; //lowcode userId standalone

const { prompts, sheet, workbook } = loadSpreadsheet(filePath);

// Clear the existing responses in the second column
clearResponses(sheet);

const socket = io('https://api.simplifypath.com', {
    path: '/socket.io',
    query: {
        userId: userId,
        source: "simplifypath_standalone_app",
        EIO: 4,
        transport: 'polling'
    },
    transports: ['polling'] // Ensure correct transport method
});



let conversation = "";
const responses = [];

function sendNextPrompt() {
    if (prompts.length > 0) {
        const nextPrompt = prompts.shift();
        socket.emit('user_reply', { message: nextPrompt, userId: userId });
        console.log(`User: ${nextPrompt}`);
        conversation += `User: ${nextPrompt}\n`;
    } else {
        fs.writeFileSync(outputFilePath, conversation);
        saveResponses(filePath, responses, sheet, workbook);
        console.log('Responses saved to the spreadsheet');
        socket.disconnect();
    }
}

socket.on('connect', () => {
    console.log('Connected to the server');
    sendNextPrompt(); // Send the first prompt
});

socket.on('connect_error', (error) => {
    console.error('Connection error:', error);
});

socket.on('disconnect', () => {
    console.log('Disconnected from the server');
});

socket.on('bot_reply', (data) => {
    console.log(`Bot: ${JSON.stringify(data)}`);
    const botMessage = data.message;
    conversation += `Bot: ${botMessage}\n\n`;
    responses.push(botMessage);

    sendNextPrompt(); // Send the next prompt
});

socket.on('error', (error) => {
    console.error('Socket error:', error);
});



function loadSpreadsheet(filePath) {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0]; // Use the first sheet
    const sheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 }); // Convert to array
    const prompts = jsonData.slice(1).map(row => row[0]).filter(prompt => prompt); // Skip empty prompts
    return { prompts, sheet, workbook, sheetName };
}

function clearResponses(sheet) {
    const range = XLSX.utils.decode_range(sheet['!ref']);
    for (let R = 1; R <= range.e.r; ++R) { // Start from row 1 to skip the header
        const cellAddress = XLSX.utils.encode_cell({ c: 1, r: R }); // Column B
        sheet[cellAddress] = { v: "" }; // Clear the cell
    }
}

function saveResponses(filePath, responses, sheet, workbook) {
    responses.forEach((response, index) => {
        const cellAddress = XLSX.utils.encode_cell({ c: 1, r: index + 1 }); // Column B, skip header
        sheet[cellAddress] = { v: response };
        if (!sheet[cellAddress].s) {
            sheet[cellAddress].s = {};
        }
        sheet[cellAddress].s.alignment = { wrapText: true }; // Enable text wrapping
    });

    // Set the width of columns
    const wscols = [
        { wch: 10 * 4.67 }, // 10 cm
        { wch: 35 * 4.67 } // 35 cm
    ];
    sheet['!cols'] = wscols;

    XLSX.writeFile(workbook, filePath);
}


// let conversation = "";
// const responses = [];

// function sendNextPrompt() {
//     if (prompts.length > 0) {
//         const nextPrompt = prompts.shift();
//         socket.emit('user_reply', { message: nextPrompt, userId: userId });
//         console.log(`User: ${nextPrompt}`);
//         conversation += `User: ${nextPrompt}\n`;
//     } else {
//         fs.writeFileSync(outputFilePath, conversation);
//         saveResponses(filePath, responses, sheet, workbook);
//         console.log('Responses saved to the spreadsheet');
//         socket.disconnect();
//     }
// }

// socket.on('connect', () => {
//     console.log('Connected to the server');
//     sendNextPrompt(); // Send the first prompt
// });

// socket.on('connect_error', (error) => {
//     console.error('Connection error:', error);
// });

// socket.on('disconnect', () => {
//     console.log('Disconnected from the server');
// });

// socket.on('bot_reply', (data) => {
//     console.log(`Bot: ${JSON.stringify(data)}`);
//     const botMessage = data.message;
//     conversation += `Bot: ${botMessage}\n\n`;
//     responses.push(botMessage);

//     sendNextPrompt(); // Send the next prompt
// });

// socket.on('error', (error) => {
//     console.error('Socket error:', error);
// });

// function loadSpreadsheet(filePath) {
//     const workbook = XLSX.readFile(filePath);
//     const sheetName = workbook.SheetNames[0]; // Use the first sheet
//     const sheet = workbook.Sheets[sheetName];
//     const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 }); // Convert to array
//     const prompts = jsonData.slice(1).map(row => row[0]).filter(prompt => prompt); // Skip empty prompts
//     return { prompts, sheet, workbook, sheetName };
// }

// function clearResponses(sheet) {
//     const range = XLSX.utils.decode_range(sheet['!ref']);
//     for (let R = 1; R <= range.e.r; ++R) { // Start from row 1 to skip the header
//         const cellAddress = XLSX.utils.encode_cell({ c: 1, r: R }); // Column B
//         sheet[cellAddress] = { v: "" }; // Clear the cell
//     }
// }

// function saveResponses(filePath, responses, sheet, workbook) {
//     responses.forEach((response, index) => {
//         const cellAddress = XLSX.utils.encode_cell({ c: 1, r: index + 1 }); // Column B, skip header
//         sheet[cellAddress] = { v: response };
//         if (!sheet[cellAddress].s) {
//             sheet[cellAddress].s = {};
//         }
//         sheet[cellAddress].s.alignment = { wrapText: true }; // Enable text wrapping
//     });

//     // Set the width of columns
//     const wscols = [
//         { wch: 10 * 4.67 }, // 10 cm
//         { wch: 35 * 4.67 } // 35 cm
//     ];
//     sheet['!cols'] = wscols;

//     XLSX.writeFile(workbook, filePath);
// }

