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
  // Your Google Sheet ID (you can find it in the URL)
  const spreadsheetId = "1eWEJdrrjIaZFF8lWXAje4KvGnc9jJtD5Jz3xjvK7y5o";
  let lineCurrently = 11;

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

    Object.keys(data).map(async (item) => {
      sheets.spreadsheets.batchUpdate({
        auth: client,
        spreadsheetId: spreadsheetId,
        resource: {
          requests: [
            {
              mergeCells: {
                range: {
                  sheetId: copiedSheetId,
                  startRowIndex: lineCurrently - 1,
                  endRowIndex: lineCurrently,
                  startColumnIndex: 0,
                  endColumnIndex: 2,
                },
                mergeType: "MERGE_ALL",
              },
            },
            {
              repeatCell: {
                range: {
                  sheetId: copiedSheetId,
                  startRowIndex: lineCurrently - 1,
                  endRowIndex: lineCurrently,
                  startColumnIndex: 0,
                  endColumnIndex: 2,
                },
                cell: {
                  userEnteredFormat: {
                    backgroundColor: {
                      red: 0.0,
                      green: 0.0,
                      blue: 0.0,
                    },
                    textFormat: {
                      foregroundColor: {
                        red: 1.0,
                        green: 1.0,
                        blue: 1.0,
                      },
                      fontSize: 8,
                    },
                  },
                },
                fields: "userEnteredFormat(backgroundColor,textFormat)",
              },
            },
          ],
        },
      });
      sheets.spreadsheets.values.update({
        auth: client,
        spreadsheetId: spreadsheetId,
        range: `A${lineCurrently}`,
        valueInputOption: "RAW",
        resource: {
          values: [[item]],
        },
      });
      lineCurrently += 1;
      sheets.spreadsheets.batchUpdate({
        auth: client,
        spreadsheetId: spreadsheetId,
        resource: {
          requests: [
            {
              mergeCells: {
                range: {
                  sheetId: copiedSheetId,
                  startRowIndex: lineCurrently - 1,
                  endRowIndex: lineCurrently,
                  startColumnIndex: 4,
                  endColumnIndex: 8,
                },
                mergeType: "MERGE_ALL",
              },
            },
            {
              mergeCells: {
                range: {
                  sheetId: copiedSheetId,
                  startRowIndex: lineCurrently - 1,
                  endRowIndex: lineCurrently,
                  startColumnIndex: 8,
                  endColumnIndex: 12,
                },
                mergeType: "MERGE_ALL",
              },
            },
            {
              mergeCells: {
                range: {
                  sheetId: copiedSheetId,
                  startRowIndex: lineCurrently - 1,
                  endRowIndex: lineCurrently,
                  startColumnIndex: 12,
                  endColumnIndex: 16,
                },
                mergeType: "MERGE_ALL",
              },
            },
            {
              mergeCells: {
                range: {
                  sheetId: copiedSheetId,
                  startRowIndex: lineCurrently - 1,
                  endRowIndex: lineCurrently,
                  startColumnIndex: 16,
                  endColumnIndex: 20,
                },
                mergeType: "MERGE_ALL",
              },
            },
            {
              mergeCells: {
                range: {
                  sheetId: copiedSheetId,
                  startRowIndex: lineCurrently - 1,
                  endRowIndex: lineCurrently,
                  startColumnIndex: 20,
                  endColumnIndex: 24,
                },
                mergeType: "MERGE_ALL",
              },
            },
            {
              mergeCells: {
                range: {
                  sheetId: copiedSheetId,
                  startRowIndex: lineCurrently - 1,
                  endRowIndex: lineCurrently,
                  startColumnIndex: 24,
                  endColumnIndex: 28,
                },
                mergeType: "MERGE_ALL",
              },
            },
            {
              mergeCells: {
                range: {
                  sheetId: copiedSheetId,
                  startRowIndex: lineCurrently - 1,
                  endRowIndex: lineCurrently,
                  startColumnIndex: 28,
                  endColumnIndex: 32,
                },
                mergeType: "MERGE_ALL",
              },
            },
            {
              mergeCells: {
                range: {
                  sheetId: copiedSheetId,
                  startRowIndex: lineCurrently - 1,
                  endRowIndex: lineCurrently,
                  startColumnIndex: 32,
                  endColumnIndex: 36,
                },
                mergeType: "MERGE_ALL",
              },
            },
            {
              mergeCells: {
                range: {
                  sheetId: copiedSheetId,
                  startRowIndex: lineCurrently - 1,
                  endRowIndex: lineCurrently,
                  startColumnIndex: 36,
                  endColumnIndex: 40,
                },
                mergeType: "MERGE_ALL",
              },
            },
            {
              mergeCells: {
                range: {
                  sheetId: copiedSheetId,
                  startRowIndex: lineCurrently - 1,
                  endRowIndex: lineCurrently,
                  startColumnIndex: 40,
                  endColumnIndex: 44,
                },
                mergeType: "MERGE_ALL",
              },
            },
            {
              mergeCells: {
                range: {
                  sheetId: copiedSheetId,
                  startRowIndex: lineCurrently - 1,
                  endRowIndex: lineCurrently,
                  startColumnIndex: 44,
                  endColumnIndex: 48,
                },
                mergeType: "MERGE_ALL",
              },
            },
            {
              mergeCells: {
                range: {
                  sheetId: copiedSheetId,
                  startRowIndex: lineCurrently - 1,
                  endRowIndex: lineCurrently,
                  startColumnIndex: 48,
                  endColumnIndex: 52,
                },
                mergeType: "MERGE_ALL",
              },
            },
          ],
        },
      });
      sheets.spreadsheets.values.update({
        auth: client,
        spreadsheetId: spreadsheetId,
        range: `A${lineCurrently}`,
        valueInputOption: "USER_ENTERED",
        resource: {
          values: [
            [
              "=Glossary!$A$6",
              "=SUM(E" + lineCurrently + ":AZ" + lineCurrently + ")",
              "=Glossary!$A$13",
              undefined,
              "=ROUND(SUM(E" + `${lineCurrently + 1}` + ":H" + `${lineCurrently + 1}` + ")/4,2)",
              undefined,
              undefined,
              undefined,
              "=ROUND(SUM(I" + `${lineCurrently + 1}` + ":L" + `${lineCurrently + 1}` + ")/4,2)",
              undefined,
              undefined,
              undefined,
              "=ROUND(SUM(M" + `${lineCurrently + 1}` + ":P" + `${lineCurrently + 1}` + ")/4,2)",
              undefined,
              undefined,
              undefined,
              "=ROUND(SUM(Q" + `${lineCurrently + 1}` + ":T" + `${lineCurrently + 1}` + ")/4,2)",
              undefined,
              undefined,
              undefined,
              "=ROUND(SUM(U" + `${lineCurrently + 1}` + ":X" + `${lineCurrently + 1}` + ")/4,2)",
              undefined,
              undefined,
              undefined,
              "=ROUND(SUM(Y" + `${lineCurrently + 1}` + ":AB" + `${lineCurrently + 1}` + ")/4,2)",
              undefined,
              undefined,
              undefined,
              "=ROUND(SUM(AC" + `${lineCurrently + 1}` + ":AF" + `${lineCurrently + 1}` + ")/4,2)",
              undefined,
              undefined,
              undefined,
              "=ROUND(SUM(AG" + `${lineCurrently + 1}` + ":AJ" + `${lineCurrently + 1}` + ")/4,2)",
              undefined,
              undefined,
              undefined,
              "=ROUND(SUM(AK" + `${lineCurrently + 1}` + ":AN" + `${lineCurrently + 1}` + ")/4,2)",
              undefined,
              undefined,
              undefined,
              "=ROUND(SUM(AO" + `${lineCurrently + 1}` + ":AR" + `${lineCurrently + 1}` + ")/4,2)",
              undefined,
              undefined,
              undefined,
              "=ROUND(SUM(AS" + `${lineCurrently + 1}` + ":AV" + `${lineCurrently + 1}` + ")/4,2)",
              undefined,
              undefined,
              undefined,
              "=ROUND(SUM(AW" + `${lineCurrently + 1}` + ":AZ" + `${lineCurrently + 1}` + ")/4,2)",
              undefined,
              undefined,
              undefined,
            ],
          ],
        },
      });
      lineCurrently += 1;
      sheets.spreadsheets.values.update({
        auth: client,
        spreadsheetId: spreadsheetId,
        range: `A${lineCurrently}`,
        valueInputOption: "USER_ENTERED",
        resource: {
          values: [
            [
              "=Glossary!$A$7",
              undefined,
              undefined,
              "=Glossary!$A$14",
              "=SUM(E" + `${lineCurrently + 2}` + ":E" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(F" + `${lineCurrently + 2}` + ":F" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(G" + `${lineCurrently + 2}` + ":G" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(H" + `${lineCurrently + 2}` + ":H" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(I" + `${lineCurrently + 2}` + ":I" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(J" + `${lineCurrently + 2}` + ":J" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(K" + `${lineCurrently + 2}` + ":K" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(L" + `${lineCurrently + 2}` + ":L" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(M" + `${lineCurrently + 2}` + ":M" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(N" + `${lineCurrently + 2}` + ":N" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(O" + `${lineCurrently + 2}` + ":O" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(P" + `${lineCurrently + 2}` + ":P" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(Q" + `${lineCurrently + 2}` + ":Q" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(R" + `${lineCurrently + 2}` + ":R" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(S" + `${lineCurrently + 2}` + ":S" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(T" + `${lineCurrently + 2}` + ":T" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(U" + `${lineCurrently + 2}` + ":U" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(V" + `${lineCurrently + 2}` + ":V" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(W" + `${lineCurrently + 2}` + ":W" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(X" + `${lineCurrently + 2}` + ":X" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(Y" + `${lineCurrently + 2}` + ":Y" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(Z" + `${lineCurrently + 2}` + ":Z" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(AA" + `${lineCurrently + 2}` + ":AA" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(AB" + `${lineCurrently + 2}` + ":AB" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(AC" + `${lineCurrently + 2}` + ":AC" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(AD" + `${lineCurrently + 2}` + ":AD" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(AE" + `${lineCurrently + 2}` + ":AE" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(AF" + `${lineCurrently + 2}` + ":AF" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(AG" + `${lineCurrently + 2}` + ":AG" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(AH" + `${lineCurrently + 2}` + ":AH" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(AI" + `${lineCurrently + 2}` + ":AI" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(AJ" + `${lineCurrently + 2}` + ":AJ" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(AK" + `${lineCurrently + 2}` + ":AK" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(AL" + `${lineCurrently + 2}` + ":AL" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(AM" + `${lineCurrently + 2}` + ":AM" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(AN" + `${lineCurrently + 2}` + ":AN" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(AO" + `${lineCurrently + 2}` + ":AO" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(AP" + `${lineCurrently + 2}` + ":AP" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(AQ" + `${lineCurrently + 2}` + ":AQ" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(AR" + `${lineCurrently + 2}` + ":AR" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(AS" + `${lineCurrently + 2}` + ":AS" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(AT" + `${lineCurrently + 2}` + ":AT" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(AU" + `${lineCurrently + 2}` + ":AU" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(AV" + `${lineCurrently + 2}` + ":AV" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(AW" + `${lineCurrently + 2}` + ":AW" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(AX" + `${lineCurrently + 2}` + ":AX" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(AY" + `${lineCurrently + 2}` + ":AY" + `${lineCurrently + 2 + data[item].length}` + ")",
              "=SUM(AZ" + `${lineCurrently + 2}` + ":AZ" + `${lineCurrently + 2 + data[item].length}` + ")",
            ],
          ],
        },
      });
      lineCurrently += 1;
      sheets.spreadsheets.batchUpdate({
        auth: client,
        spreadsheetId: spreadsheetId,
        resource: {
          requests: [
            {
              mergeCells: {
                range: {
                  sheetId: copiedSheetId,
                  startRowIndex: lineCurrently - 1,
                  endRowIndex: lineCurrently,
                  startColumnIndex: 0,
                  endColumnIndex: 2,
                },
                mergeType: "MERGE_ALL",
              },
            },
          ],
        },
      });
      sheets.spreadsheets.values.update({
        auth: client,
        spreadsheetId: spreadsheetId,
        range: `A${lineCurrently}`,
        valueInputOption: "RAW",
        resource: {
          values: [["Resource Allocation"]],
        },
      });
      lineCurrently += 1;

      // For Resource Allocation
      data[item].map((item, index) => {
        const newItem = [...item];
        newItem[0] = newItem[2];
        newItem[2] = "=D" + lineCurrently + "/4";
        newItem[3] = "=SUM(E" + `${lineCurrently}` + ":AZ" + `${lineCurrently}` + ")";
        sheets.spreadsheets.values.update({
          auth: client,
          spreadsheetId: spreadsheetId,
          range: `A${lineCurrently}`,
          valueInputOption: "USER_ENTERED",
          resource: {
            values: [newItem],
          },
        });
        lineCurrently += 1;
      });
      lineCurrently += 1;
    });

    sheets.spreadsheets.batchUpdate({
      auth: client,
      spreadsheetId: spreadsheetId,
      resource: {
        requests: [
          {
            repeatCell: {
              range: {
                sheetId: copiedSheetId,
                startRowIndex: 0,
                endRowIndex: lineCurrently - 1,
                startColumnIndex: 0,
                endColumnIndex: 52,
              },
              cell: {
                userEnteredFormat: {
                  borders: {
                    top: {
                      style: "DOTTED",
                      color: {
                        red: 0.0,
                        green: 0.0,
                        blue: 0.0,
                      },
                    },
                    bottom: {
                      style: "DOTTED",
                      color: {
                        red: 0.0,
                        green: 0.0,
                        blue: 0.0,
                      },
                    },
                    left: {
                      style: "DOTTED",
                      color: {
                        red: 0.0,
                        green: 0.0,
                        blue: 0.0,
                      },
                    },
                    right: {
                      style: "DOTTED",
                      color: {
                        red: 0.0,
                        green: 0.0,
                        blue: 0.0,
                      },
                    },
                  },
                },
              },
              fields: "userEnteredFormat.borders",
            },
          },
        ],
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
