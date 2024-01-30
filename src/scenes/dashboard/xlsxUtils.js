import { initializeApp } from "firebase/app";
import { getAnalytics } from "firebase/analytics";
import { getFirestore, collection, addDoc } from 'firebase/firestore';

// TODO: Add SDKs for Firebase products that you want to use
// https://firebase.google.com/docs/web/setup#available-libraries

// Your web app's Firebase configuration
// For Firebase JS SDK v7.20.0 and later, measurementId is optional
const firebaseConfig = {
  apiKey: "AIzaSyBUImVg0vjCqNWh-8JKETch9XatQ2bj19o",
  authDomain: "pln-app-952f5.firebaseapp.com",
  projectId: "pln-app-952f5",
  storageBucket: "pln-app-952f5.appspot.com",
  messagingSenderId: "555837575105",
  appId: "1:555837575105:web:1311558b34d8e7040ea940",
  measurementId: "G-KQ2XNCG6QY"
};

// Initialize Firebase
// Initialize Firestore
const firestore = getFirestore(app);
const app = initializeApp(firebaseConfig);
const analytics = getAnalytics(app);const XLSX = require('xlsx');
const fs = require('fs');

function check(value) {
    try {
        parseFloat(value);
        return true;
    } catch (error) {
        try {
            return parseInt(value) >= 1 && parseInt(value) <= 12;
        } catch (error) {
            return false;
        }
    }
}

async function xlsxToJson(filename, range) {
    try {
        const workbook = XLSX.readFile(filename);

        const data = {};

        workbook.SheetNames.forEach(sheetName => {
            const currentSheet = workbook.Sheets[sheetName];
            let currentRange = range;

            for (let rowIndex = currentRange.startRow; rowIndex <= currentRange.endRow; rowIndex++) {
                const row = XLSX.utils.sheet_to_json(currentSheet, { header: 1, range: `${currentRange.startCol}${rowIndex}:${currentRange.endCol}${rowIndex}` })[0];

                // kurang HPL & RPT - RCT
                if (sheetName === 'H P L' || sheetName === "RPT-RCT") {
                    break;
                }

                if (row.some(cellValue => cellValue !== null && cellValue !== '' && check(cellValue))) {
                    const cleanedData = row.map(value => (value !== null && check(value) ? value : ''));

                    const sheetKey = sheetName.toLowerCase().replace(/ /g, '_'); 
                    if (!data[sheetKey]) {
                        data[sheetKey] = {};
                    }

                    const setKey = `row${rowIndex - currentRange.startRow + 1}`;
                    data[sheetKey][setKey] = {
                        values: cleanedData
                    };
                }
            }
        });

        const docRef = await addDoc(collection(firestore,'2024'), {
            data,
        });
        console.log('Document written with ID: ', docRef.id);

        return data;
    } catch (error) {
        console.error(`Error: ${error.message}`);
        return null;
    }
}

const targetRange = {
    startCol: 'E',
    endCol: 'P',
    startRow: 6,
    endRow: 29,
    groupCol: 0
};

const jsonOutput = xlsxToJson('Kinerja SARPP - 2023.12.-magang.xlsx', targetRange);

if (jsonOutput !== null) {
    console.log('JSON Output:');
    console.log(jsonOutput);
    console.log('Output saved as output_js.json');
}

export default xlsxToJson;