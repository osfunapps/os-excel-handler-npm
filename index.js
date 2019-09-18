let Excel = require('exceljs');


/** Determined this value based upon experimentation */
PIXELS_PER_EXCEL_WIDTH_UNIT = 6;
const self = module.exports = {

    /**
     * Will merge a few cells together.
     *
     * @param sheet -> the sheet you work on
     * @param cellIdStart -> like A1
     * @param cellIdEnd -> like A2
     */
    mergeCells(sheet, cellIdStart, cellIdEnd) {
        sheet.mergeCells(cellIdStart + ':' + cellIdEnd);
    },

    /**
     * Will set a default font to all of the empty cells in the sheet.
     */
    setFontInAllEmptyCells(sheet, name = 'Arial', size = 12, bold = false, italic = false) {
        sheet.columns.forEach(function (row) {
            row.style = {font: {name: name, size: size, bold: bold, italic: italic}};
        });
    },

    /**
     * Will fit the column width to be in the size of the biggest line in the column
     */
    fitColumnWidthToText(sheet, colLettersArr = []) {
        let maxColumnLength = 0;
        for (let i = 0; i < colLettersArr.length; i++) {
            let colLetter = colLettersArr[i];

            sheet.getColumn(colLetter).eachCell((cell) => {
                if (typeof cell.value === 'string') {

                    const fontSize = cell.font && cell.font.size ? cell.font.size : 11;
                    let pixelWidth = require('string-pixel-width');
                    const cellWidth = pixelWidth(cell.value, {size: fontSize});

                    maxColumnLength = Math.max(maxColumnLength, cellWidth)
                }
            });

            sheet.getColumn(colLetter).width = maxColumnLength / PIXELS_PER_EXCEL_WIDTH_UNIT + 1
            maxColumnLength = 0
        }
    },

    /**
     * Also called wrap text. Will set the height of the line to be in the size of the cell content.
     */
    wrapText(sheet, cellId) {
        if (sheet.getCell(cellId).alignment === undefined) {
            sheet.getCell(cellId).alignment = {};
        }
        sheet.getCell(cellId).alignment.wrapText = true
    },

    /**
     * Will align the text in the cell to be in the middle and center
     */
    alignCellCenter(sheet, cellId) {
        let cell = sheet.getCell(cellId)
        alignCenter(cell)
    },

    /**
     * Will align the text in the element to be in the middle and center
     */
    alignCenter(ele) {
      alignCenter(ele)
    },

    /**
     * Will change the row height.
     * NOTICE: the default raw height, if didn't changed, is 15.
     *
     * @param row -> the row in question
     * @param newRowHeight -> the new height of the row
     */
    setRowHeight(row, newRowHeight) {
      row.height = newRowHeight
    },


    /**
     * Will read a value from a given cell identifier
     *
     * @param sheet -> the current sheet obj
     * @param cellId -> row and column identifier. Like A4
     */
    readValue(sheet, cellId) {
        return sheet.getCell(cellId).header
    },

    /**
     * Will set a value in a given cell id
     *
     * @param sheet -> the current sheet obj
     * @param cellId -> row and column identifier. Like A4
     * @param value -> the value you would like to write
     * @param fontName -> the name of the font in the cell
     * @param fontSize -> the font size in the cell
     * @param bold -> font bold toggle
     * @param italic -> font italic toggle
     */
    setValue(sheet, cellId, value, fontName = 'Arial', fontSize = 12, bold = false, italic = false) {
        sheet.getCell(cellId).value = value;
        sheet.getCell(cellId).style.font = {
            name: fontName,
            size: fontSize,
            bold: bold,
            italic: italic
        };
    },

    /**
     * Will add a sheet to the workbook
     */
    createSheet(wb, sheetName, rtl = false) {
        let sheet = wb.addWorksheet(sheetName);
        if (rtl) {
            sheet.views = [{rightToLeft: true}];
        }
        return sheet
    },

    /**
     * Will create a new workbook to start the work on
     */
    createWorkbook() {
        return new Excel.Workbook();
    },

    /**
     *  Will save the workbook in a given path
     */
    saveWorkbook: async function (wb, workbookPath) {
        await wb.xlsx.writeFile(workbookPath)
            .then(function () {
                // done
            });
    },

    /**
     * Will set the background color of a given cell
     */
    setCellBackgroundColor(sheet, cellId, hexColor) {
        let cell = sheet.getCell(cellId)
        setBackgroundColor(cell, hexColor)
    },


    /**
     * Will set the background color of a given element
     */
    setElementBackgroundColor(ele, hexColor) {
        setBackgroundColor(ele, hexColor)
    },

    /**
     * Will set the text color of the text in a given cell
     */
    setCellTextColor(sheet, cellId, hexColor) {
        sheet.getCell(cellId).font = {
            color: {argb: hexColor},
        };
    },

    /**
     * Will return a row
     */
    getRow(sheet, row) {
        return sheet.getRow(row)
    },

    /**
     * Will return a column
     */
    getColumn(sheet, column) {
        return sheet.getColumn(column)
    },

    /**
     * Will change the entire style of an element
     */
    setEleStyle(ele,
                fontName = 'Arial',
                fontSize = 12,
                bold = false,
                italic = false,
                backgroundColor = null,
                fontColor = null) {
        ele.eachCell((cell) => {
            cell.style.font = {
                    name: fontName,
                    size: fontSize,
                    bold: bold,
                    italic: italic
            };
            if (fontColor !== null) {
                cell.style.font.color = {argb: fontColor}
            }
        });
        if (backgroundColor !== null) {
            self.setElementBackgroundColor(ele, backgroundColor)
        }
    }
};

function setBackgroundColor(ele, hexColor) {
    ele.fill = {
        type: 'gradient',
        gradient: 'angle',
        degree: 0,
        stops: [
            {position: 0, color: {argb: hexColor}},
            {position: 1, color: {argb: hexColor}},
        ]
    };
}

function alignCenter(ele) {
    if (ele.alignment === undefined) {
        ele.alignment = {};
    }
    ele.alignment.vertical = 'middle';
    ele.alignment.horizontal = 'center';
}