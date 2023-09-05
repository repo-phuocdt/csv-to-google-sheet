const fs = require("fs");
const { parse } = require("csv-parse");
const { groupBy } = require("lodash"); // Import lodash
const { google } = require("googleapis");
const sheets = google.sheets("v4");

const credentials = require("./credentials.json");

const client = new google.auth.JWT(credentials.client_email, null, credentials.private_key, [
  "https://www.googleapis.com/auth/spreadsheets",
]);

async function createSheetAndWriteData(data) {
  console.log(data);
  // Your Google Sheet ID (you can find it in the URL)
  const spreadsheetId = "1eWEJdrrjIaZFF8lWXAje4KvGnc9jJtD5Jz3xjvK7y5o";

  const currentDate = new Date();
  // Get the year, month, day, hours, minutes and seconds from the current date
  const year = currentDate.getFullYear();
  const month = String(currentDate.getMonth() + 1).padStart(2, "0");
  const day = String(currentDate.getDate()).padStart(2, "0");
  const hours = String(currentDate.getHours()).padStart(2, "0");
  const minutes = String(currentDate.getMinutes()).padStart(2, "0");
  const seconds = String(currentDate.getSeconds()).padStart(2, "0");

  // Construct the formatted date string
  const newSheetTitle = `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;

  try {
    const response = await sheets.spreadsheets.get({
      auth: client,
      spreadsheetId: spreadsheetId,
    });
    const sheetsList = response.data.sheets || [];
    const sheetIdTemplate = sheetsList.find((item) => item?.properties?.title === "Template")?.properties.sheetId;

    const copyResponse = await sheets.spreadsheets.sheets.copyTo({
      auth: client,
      spreadsheetId: spreadsheetId,
      sheetId: sheetIdTemplate,
      resource: {
        destinationSpreadsheetId: spreadsheetId,
      },
    });

    const copiedSheetId = copyResponse.data.sheetId;

    await sheets.spreadsheets.batchUpdate({
      auth: client,
      spreadsheetId: spreadsheetId,
      resource: {
        requests: [
          {
            updateSheetProperties: {
              properties: {
                sheetId: copiedSheetId,
                title: newSheetTitle,
              },
              fields: "title",
            },
          },
          {
            updateSheetProperties: {
              properties: {
                sheetId: copiedSheetId,
                index: 0,
              },
              fields: "index",
            },
          },
        ],
      },
    });

    // Data to write
    const values = [
      ["Value1", "Value2", "Value33333"],
      ["Value4", "Value5", "Value6"],
    ];

    // Write data to the Google Sheet
    await sheets.spreadsheets.values.update({
      auth: client,
      spreadsheetId: spreadsheetId,
      range: "A1",
      valueInputOption: "RAW",
      resource: {
        values: values,
      },
    });
  } catch (error) {
    console.error("An error occurred:", error);
  }
}

async function handleUploadCSVToGoogleSheet() {
  // Load CSV data
  const data = [];
  fs.createReadStream("./resource-list.csv")
    .pipe(parse({ delimiter: ",", from_line: 3 }))
    .on("data", (row) => data.push(row))
    .on("end", () => {
      const groupedData = groupBy(data, (subArray) => subArray[0]);

      client.authorize(async function (err) {
        if (err) {
          console.error("Authentication failed:", err);
          return;
        }
        createSheetAndWriteData(groupedData);
      });
    });
}

handleUploadCSVToGoogleSheet();
