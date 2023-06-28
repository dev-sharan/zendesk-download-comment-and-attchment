const xlsx = require("xlsx");
const workbook = xlsx.readFile("Closed Ticket Scenarios.xlsx");
const axios = require("axios");
const { encode } = require("base-64");
const XlsxPopulate = require("xlsx-populate");
const progress = require("progress");
const fs = require("fs");

const sheetNames = workbook.SheetNames;

async function makeApiCall(url, path, method, username, password) {
  const authorization = `Basic ${encode(`${username}:${password}`)}`;

  const config = {
    headers: {
      Authorization: authorization,
    },
  };

  try {
    const response = await axios.get(`${url}${path}`, config);
    return response.data;
  } catch (error) {
    throw error;
  }
}

(async () => {
  const username = "xxxxxxxxxxx";
  const password = "xxxxxxxxxxx";

  let dataexp = [];

  const progressBar = new progress(
    "Fetching ticket comments [:bar] :current/:total :percent :etas",
    {
      total: sheetNames.length,
      width: 20,
    }
  );

  for (const sheetName of sheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const sheetData = xlsx.utils.sheet_to_json(sheet);

    for (const row of sheetData) {
      const ticketId = row["Ticket ID"];
      const url = "https://xxxxxxxxxxxxxx.zendesk.com/api/v2/";
      const path = `tickets/${ticketId}/comments?include=ticket_event.json`;
      const method = "GET";

      try {
        const data = await makeApiCall(url, path, method, username, password);
        dataexp.push({
          ticketId: ticketId,
          comments: data.comments,
        });

        // Save attachment files and update comment body with file reference
        let attachmentCounter = 1;
        for (const dt of dataexp) {
          for (const [index, comment] of dt.comments.entries()) {
            if (comment.attachments.length > 0) {
              const attachments = comment.attachments;

                    // Save attachment files locally
                    for (let i = 0; i < attachments.length; i++) {
                      const attachment = attachments[i];
                      const attachmentData = await axios.get(
                        attachment.content_url,
                        {
                          responseType: "arraybuffer",
                          headers: {
                            Authorization: `Basic ${encode(
                              `${username}:${password}`
                            )}`,
                          },
                        }
                      );

                      const directoryPath = `./attch/${dt.ticketId}/${index+1}`;
                      fs.mkdirSync(directoryPath,{recursive: true}, (error) => {
                        if (error) {
                          console.error('Error creating directory:', error);
                        } else {
                          console.log('Directory created successfully.');
                        }
                      });
                      
                      const attachmentFilename = `ticketno_${dt.ticketId}_comment_${index+1}_count_${attachmentCounter}_file_${attachment.file_name}`;

                      const filePath = `${directoryPath}/${attachmentFilename}`;
                      fs.writeFileSync(filePath, attachmentData.data);

                      // Update comment body with file reference
                      comment.body += `\n \nAttachment::: ${attachmentFilename}`;

                      attachmentCounter++;
                      if(i === attachments.length - 1) {
                        attachmentCounter = 1;
                      }
                    }
            }
        }
              
            
          
        }
      } catch (error) {
        console.error(error);
      }
    }

    
  }

  // Exporting into Excel format
  XlsxPopulate.fromBlankAsync()
    .then((workbook) => {
      const sheet = workbook.sheet(0);
      const maxColumnLength = Math.max(
        ...dataexp.map((row) => row.comments.length)
      );
      const columnHeaders = ["Ticket Id"];
      for (let i = 1; i <= maxColumnLength; i++) {
          columnHeaders.push(`Comment ${i}`);
          columnHeaders.push(`Attachment ${i}`);
      }
      columnHeaders.forEach((header, index) => {
        sheet.cell(1, index + 1).value(header);
      });

      dataexp.forEach((row, rowIndex) => {
        sheet.cell(rowIndex + 2, 1).value(row.ticketId); // Add ticketId to the first column
        let atchisthere = []
        for(let i = 0; i < row.comments.length*2 ;) {
          const columnIndex = i + 2;
          let dt = i/2;
          let str = row.comments[dt].body+'';
          str.includes('Attachment:::') ? atchisthere.push(columnIndex+1):'';
          sheet.cell(rowIndex + 2, columnIndex).value(str);
          i = i+2;
        }
        
        for(let i = 1; i < row.comments.length*2 ;) {
          const columnIndex = i + 2;
          let url = `./attch/${row.ticketId}/${i/2+.5}`
          if(atchisthere.includes(columnIndex)) {
            sheet
            .cell(rowIndex + 2, columnIndex)
            .value('Attachments')
            .style({ fontColor: "0563c1", underline: true })
            .hyperlink(url);
          }
          
            
          i = i+2;
        }

      });
      progressBar.tick();
      return workbook.toFileAsync("output.xlsx");
    })
    .then(() => {
      console.log("Spreadsheet created successfully.");
    })
    .catch((error) => {
      console.error("Error:", error);
    });
})();
