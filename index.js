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
         * NOTICE: remember to do this at the end, after you populated all of the cells with values
         *
         * @param sheet -> your excel spread sheet
         * @param colLettersArr -> an array of all of the columns you wish to change width
         * @param ignoredRows -> an array of all of the rows you wouldn't like to consider while changing
         * the width
         */
        fitColumnWidthToText(sheet, colLettersArr = [], ignoredRows = []) {
            for (let i = 0; i < colLettersArr.length; i++) {
                let maxColumnLength = 0;
                let colLetter = colLettersArr[i];
                sheet.getColumn(colLetter).eachCell((cell) => {
                    if (!ignoredRows.includes(cell.row)) {
                        if (typeof cell.value === 'string') {

                            const fontSize = cell.font && cell.font.size ? cell.font.size : 11;
                            let pixelWidth = require('string-pixel-width');
                            const cellWidth = pixelWidth(cell.value, {size: fontSize});

                            maxColumnLength = Math.max(maxColumnLength, cellWidth)
                        }
                    }
                });

                sheet.getColumn(colLetter).width = maxColumnLength / PIXELS_PER_EXCEL_WIDTH_UNIT + 1;
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
         * * NOTICE: if you're trying to center align a merged cell, just pick one of the merged cells
         * and put it here
         */
        alignCellCenter(sheet, cellId) {
            let cell = sheet.getCell(cellId);
            alignCenter(cell)
        },

        /**
         * Will align the text in the element to be in the middle and center.
         * NOTICE: if you're trying to center align a merged cell, just pick one of the merged cells
         * and put it here
         */
        alignCenter(ele) {
            alignCenter(ele)
        },

        /**
         * Will return a bunch of cells, from range and from specific locations
         *
         * @param sheet -> your current excel sheet
         * @param range -> 'A1:A5', for example
         * @param locationsArr -> ['A1', 'C66', 'F7'], for example
         */
        getCells(sheet, range = null, locationsArr = null) {
            let cellsStrArr = [];
            if (range !== null) {
                let rangeArr = range.split(':');
                let letterStart = rangeArr[0].match(/[A-Za-z]+/g)[0];
                let letterStartASCII = letterStart.charCodeAt(0);
                let rangeStart = parseInt(rangeArr[0].match(/[^A-Za-z]+/g)[0]);
                let letterEnd = rangeArr[1].match(/[A-Za-z]+/g)[0];
                let letterEndASCII = letterEnd.charCodeAt(0);
                let rangeEnd = parseInt(rangeArr[1].match(/[^A-Za-z]+/g)[0]);

                if (letterStart !== letterEnd) {
                    for (let i = letterStartASCII; i < letterEndASCII + 1; i++) {
                        let currLetter = String.fromCharCode(i);
                        for (let j = rangeStart; j < rangeEnd + 1; j++) {
                            let ele = currLetter + j;
                            cellsStrArr.push(ele)
                        }
                    }

                } else {
                    for (let i = rangeStart; i < rangeEnd + 1; i++) {
                        cellsStrArr.push(letterStart + i)
                    }
                }

            }

            if (locationsArr !== null) {
                for (let i = 0; i < locationsArr.length; i++) {
                    cellsStrArr.push(locationsArr[i])
                }
            }

            cellsStrArr = cellsStrArr.filter(function (elem, pos) {
                return cellsStrArr.indexOf(elem) === pos;
            });

            let cellsArr = [];
            for (let i = 0; i < cellsStrArr.length; i++) {
                cellsArr.push(self.getCell(sheet, cellsStrArr[i]))
            }

            return cellsArr;

        },

        /**
         * Will return a specific cell from the worksheet
         */
        getCell(sheet, cellId) {
            return sheet.getCell(cellId)
        }
        ,

        /**
         * Will change the row height.
         * NOTICE: the default raw height, if didn't changed, is 15.
         *
         * @param row -> the row in question
         * @param newRowHeight -> the new height of the row
         */
        setRowHeight(row, newRowHeight) {
            row.height = newRowHeight
        }
        ,


        /**
         * Will read a value from a given cell identifier
         *
         * @param sheet -> the current sheet obj
         * @param cellId -> row and column identifier. Like A4
         */
        readValue(sheet, cellId) {
            return sheet.getCell(cellId).header
        }
        ,

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
        }
        ,

        /**
         * Will add a sheet to the workbook
         */
        createSheet(wb, sheetName, rtl = false) {
            let sheet = wb.addWorksheet(sheetName);
            if (rtl) {
                sheet.views = [{rightToLeft: true}];
            }
            return sheet
        }
        ,

        /**
         * Will create a new workbook to start the work on
         */
        createWorkbook() {
            return new Excel.Workbook();
        }
        ,

        /**
         *  Will save the workbook in a given path
         */
        saveWorkbook: async function (wb, workbookPath) {
            await wb.xlsx.writeFile(workbookPath)
                .then(function () {
                    // done
                });
        }
        ,

        /**
         * Will set the background color of a given cell
         */
        setCellBackgroundColor(sheet, cellId, hexColor) {
            let cell = sheet.getCell(cellId);
            setBackgroundColor(cell, hexColor)
        }
        ,


        /**
         * Will set the background color of a given element
         */
        setElementBackgroundColor(ele, hexColor) {
            setBackgroundColor(ele, hexColor)
        }
        ,

        /**
         * Will set the text color of the text in a given cell
         */
        setCellTextColor(sheet, cellId, hexColor) {
            sheet.getCell(cellId).font = {
                color: {argb: hexColor},
            };
        }
        ,

        /**
         * Will return a row
         */
        getRow(sheet, row) {
            return sheet.getRow(row)
        }
        ,

        /**
         * Will return a column
         */
        getColumn(sheet, column) {
            return sheet.getColumn(column)
        }
        ,

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
    }
;

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