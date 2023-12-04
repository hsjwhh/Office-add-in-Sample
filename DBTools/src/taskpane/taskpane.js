/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("splitCell").onclick = splitCell;
    autoCheck();
  }
});

export async function autoCheck() {
  await Excel.run(async (context) => {
    const workBook = context.workbook;
    workBook.load("name");
    await context.sync();

    // console.log(workBook.name);
    if (workBook.name.indexOf("共享平台") > -1) {
      document.getElementById("lablerun").hidden = false;
      document.getElementById("lablemsg").hidden = true;
      let sharkFolderButton = document.getElementById("sharefolder");
      sharkFolderButton.hidden = false;
      sharkFolderButton.disabled = false;
      sharkFolderButton.onclick = shareFolder;
    }
  });
}

export async function splitCell() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const dataSheet = context.workbook.worksheets.getActiveWorksheet();
      const lastRow = dataSheet.getRange("A1").getExtendedRange("Down").getLastRow();
      const rangeTitle = dataSheet.getRange("A1").getEntireRow();
      let sheets = context.workbook.worksheets;
      let oldRow = 0;
      let sampleTxt = "";
      let currentTxt = "";

      // Read the range ...
      lastRow.load("rowIndex");

      // Update the fill color
      // range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The Last rowindex was ${lastRow.rowIndex.toString()}.`);

      for (let rIndex = 1; rIndex < lastRow.rowIndex + 2; rIndex++) {
        let rangeA1 = dataSheet.getCell(rIndex, 0);
        rangeA1.load("text");
        await context.sync();
        currentTxt = rangeA1.text.toString();
        if (currentTxt !== sampleTxt) {
          if (oldRow !== 0) {
            let rangeCopy = rangeA1.getRowsAbove(rIndex - oldRow).getEntireRow();
            context.workbook.worksheets.getItem(sampleTxt).getRange("A2").copyFrom(rangeCopy);
          }
          if (currentTxt !== "") {
            oldRow = rIndex;
            sheets.add(currentTxt);
            sampleTxt = currentTxt;
            context.workbook.worksheets.getItem(currentTxt).getRange("A1").copyFrom(rangeTitle);
            // console.log(`The new sample text was ${sampleTxt}. OldRow was ${oldRow.toString()}.`);
          }
        }
      }

      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function shareFolder() {
  try {
    await Excel.run(async (context) => {
      const dataSheet = context.workbook.worksheets.getActiveWorksheet();
      const lastRow = dataSheet.getRange("A5").getExtendedRange("Down").getLastRow();
      let newDirectory = "";
      let newFolderName = "";
      let shareFolderLink = "";

      let folderNameRoot = "";
      let folderNameOne = "";
      let folderNameTwo = "";
      let folderNameThree = "";
      let folderNameFour = "";
      let folderNameFive = "";
      let folderNameSix = "";

      lastRow.load("rowIndex");
      await context.sync();
      console.log(`The Last rowindex was ${lastRow.rowIndex.toString()}.`);

      for (let rIndex = 4; rIndex < lastRow.rowIndex + 1; rIndex++) {
        let rangeDirectory = dataSheet.getCell(rIndex, 2);
        let rangeFolderName = dataSheet.getCell(rIndex, 3);
        let rangeShareFolderLink = dataSheet.getCell(rIndex, 10);
        rangeDirectory.load("text");
        rangeFolderName.load("text");
        rangeShareFolderLink.load("values");
        await context.sync();

        newDirectory = rangeDirectory.text.toString();
        newFolderName = rangeFolderName.text.toString();

        switch (newDirectory) {
          case "根目录":
            folderNameRoot = newFolderName;
            shareFolderLink = `\\\\XXXX.XXX.XXX.XXX\\共享平台\\${folderNameRoot}\\`;
            break;
          case "一级":
            folderNameOne = newFolderName;
            shareFolderLink = `\\\\XXXX.XXX.XXX.XXX\\共享平台\\${folderNameRoot}\\${folderNameOne}\\`;
            break;
          case "二级":
            folderNameTwo = newFolderName;
            shareFolderLink = `\\\\XXXX.XXX.XXX.XXX\\共享平台\\${folderNameRoot}\\${folderNameOne}\\${folderNameTwo}\\`;
            break;
          case "三级":
            folderNameThree = newFolderName;
            shareFolderLink = `\\\\XXXX.XXX.XXX.XXX\\共享平台\\${folderNameRoot}\\${folderNameOne}\\${folderNameTwo}\\${folderNameThree}\\`;
            break;
          case "四级":
            folderNameFour = newFolderName;
            shareFolderLink = `\\\\XXXX.XXX.XXX.XXX\\共享平台\\${folderNameRoot}\\${folderNameOne}\\${folderNameTwo}\\${folderNameThree}\\${folderNameFour}\\`;
            break;
          case "五级":
            folderNameFive = newFolderName;
            shareFolderLink = `\\\\XXXX.XXX.XXX.XXX\\共享平台\\${folderNameRoot}\\${folderNameOne}\\${folderNameTwo}\\${folderNameThree}\\${folderNameFour}\\${folderNameFive}\\`;
            break;
          case "六级":
            folderNameSix = newFolderName;
            shareFolderLink = `\\\\XXXX.XXX.XXX.XXX\\共享平台\\${folderNameRoot}\\${folderNameOne}\\${folderNameTwo}\\${folderNameThree}\\${folderNameFour}\\${folderNameFive}\\${folderNameSix}\\`;
            break;
        }
        rangeShareFolderLink.values = [[shareFolderLink.toString()]];
      }
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}