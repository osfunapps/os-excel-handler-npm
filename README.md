Introduction
------------

This module contains functions to deal with excel files.

## Installation
Install via npm:
    
    npm i os-excel-handler


## Usage       
Require excel handler:
        
    var eh = require("os-excel-handler")

## Functions and signatures:

    /**
     * Will merge a few cells together.
     *
     * @param sheet -> the sheet you work on
     * @param cellIdStart -> like A1
     * @param cellIdEnd -> like A2
     */
    mergeCells(sheet, cellIdStart, cellIdEnd)

    /**
     * Will set a default font to all of the empty cells in the sheet.
     */
    setFontInAllEmptyCells(sheet, name='Arial', size=12, bold=false, italic=false)

    /**
     * Will fit the column width to be in the size of the biggest line in the column
     */
    fitColumnWidthToText(sheet, colLetter)

    /**
     * Also called wrap text. Will set the height of the line to be in the size of the cell content.
     */
    wrapText(sheet, cellId) 

    /**
     * Will align the text in the cell to be in the middle and center
     */
    alignCenter(sheet, cellId)

    /**
     * Will read a value from a given cell identifier
     *
     * @param sheet -> the current sheet obj
     * @param cellId -> row and column identifier. Like A4
     */
    readValue(sheet, cellId)

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
    setValue(sheet, cellId, value, fontName='Arial', fontSize=12, bold=false, italic=false)

    /**
     * Will add a sheet to the workbook
     */
    createSheet(wb, sheetName, rtl = false)

    /**
     * Will create a new workbook to start the work on
     */
    createWorkbook()

    /**
     *  Will save the workbook in a given path
     */
    saveWorkbook: async function (wb, workbookPath)

    /**
     * Will set the background color of a given cell
     */
    setCellBackgroundColor(sheet, cellId, hexColor)

    /**
     * Will set the text color of the text in a given cell
     */
    setCellTextColor(sheet, cellId, hexColor)

And more...


## Links -> see more tools
* [os-tools-npm](https://github.com/osfunapps/os-tools-npm) -> This module contains fundamental functions to implement in an npm project
* [os-file-handler-npm](https://github.com/osfunapps/os-file-handler-npm) -> This module contains fundamental files manipulation functions to implement in an npm project
* [os-file-stream-handler-npm](https://github.com/osfunapps/os-file-stream-handler-npm) -> This module contains read/write and more advanced operations on files
* [os-xml-handler-npm](https://github.com/osfunapps/os-xml-handler-npm) -> This module will build, read and manipulate an xml file. Other handy stuff is also available, like search for specific nodes

[GitHub - osfunappsapps](https://github.com/osfunapps)



## Licence
ISC