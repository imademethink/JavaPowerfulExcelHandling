package com.PowerfulExcelWriting;
import java.io.*;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.*;
import org.junit.Assert;

public class Util01_XLCreation {

    public static void CreateXL(String xlPath, String sheetName, boolean overWriteExisting) throws Exception{

        if (! overWriteExisting && new File(xlPath).exists()){
            Assert.fail("\n\nError: XL file exists so can not over write");
        }

        Workbook workbook               = new XSSFWorkbook();
        if(null== sheetName) {sheetName = "Sheet1";}
        workbook.createSheet(sheetName);
        FileOutputStream fileOutStr      = new FileOutputStream(xlPath);
        workbook.write(fileOutStr);
        workbook.close();
        fileOutStr.close();
    }

    public static void CopyXL(String srcPath, String dstPath) throws Exception{

        FileInputStream  fileInStr  = null;
        FileOutputStream fileOutStr = null;
        try {
            fileInStr     = new FileInputStream(srcPath);
            fileOutStr    = new FileOutputStream(dstPath);
            byte[] buffer = new byte[1024];
            int length = 0;
            while ((length = fileInStr.read(buffer)) > 0) {
                fileOutStr.write(buffer, 0, length);
            }
        } finally {
            fileInStr.close();
            fileOutStr.close();
        }
    }

    public static void DeleteXL(String xlPath) throws Exception{
        File xlFile = new File(xlPath);
        if (xlFile.exists()){
            if( ! xlFile.delete()){
                Assert.fail("\n\nError: XL file deletion failed");
            }
        }else{
            Assert.fail("\n\nError: XL file does not exist");
        }
    }

    public static void AddBlankWorksheetAfterIndex(String xlPath, int index, String sheetName) throws Exception{

        FileInputStream fileInStr    = new FileInputStream(new File(xlPath));
        XSSFWorkbook workbook        = new XSSFWorkbook(fileInStr);

        if(index > workbook.getNumberOfSheets()){
            Assert.fail("Error: Existing sheet count " + workbook.getNumberOfSheets() + " is less. So can not add more sheets");
        }

        workbook.createSheet(sheetName);
//        workbook.setSheetOrder("Sheet"+Integer.toString(index),2);
        fileInStr.close();
        FileOutputStream fileOutStr  = new FileOutputStream(xlPath);
        workbook.write(fileOutStr);
        workbook.close();
        fileOutStr.close();
    }

    public static void CloneWorksheet(String xlPath, int oldSheetIndex, int newSheetIndex) throws Exception{

        FileInputStream fileInStr    = new FileInputStream(new File(xlPath));
        XSSFWorkbook workbook        = new XSSFWorkbook(fileInStr);

        if(oldSheetIndex >= workbook.getNumberOfSheets()){
            Assert.fail("Error: Existing sheet count is less. So can not add more sheets");
        }
        if(newSheetIndex > workbook.getNumberOfSheets()){
            Assert.fail("Error: New sheet count is more. So can not add more sheets");
        }

        workbook.setActiveSheet(oldSheetIndex);
        String oldSheetName          = workbook.getSheetName(oldSheetIndex);
        String newSheetName          = "new" + oldSheetName;
        workbook.cloneSheet(oldSheetIndex,"new" + oldSheetName);
        workbook.setSheetOrder(newSheetName,newSheetIndex);
        fileInStr.close();

        FileOutputStream fileOutStr  = new FileOutputStream(xlPath);
        workbook.write(fileOutStr);
        workbook.close();
        fileOutStr.close();
    }

    public static void CopyWorksheetFromTo(String fromXLPath, int fromSheetIndex,
                                            String toXLPath,   int toSheetIndex) throws Exception{
        if (! new File(fromXLPath).exists()){
            Assert.fail("\n\nError: Source excel file does not exist");
        }
        if (! new File(toXLPath).exists()){
            Assert.fail("\n\nError: Destination excel file does not exist");
        }

        FileInputStream xlSrc  = new FileInputStream(new File(fromXLPath));
        Workbook workbookSrc   = new XSSFWorkbook(xlSrc);
        if(fromSheetIndex >= workbookSrc.getNumberOfSheets()){
            Assert.fail("Error: Source XL file index is less than required index");
        }

        FileOutputStream xlDst   = new FileOutputStream(new File(toXLPath));
        Workbook workbookDst     = new XSSFWorkbook();
        if(toSheetIndex > workbookDst.getNumberOfSheets()){
            Assert.fail("Error: Destination XL file index is less than required index");
        }


//        CellStyle newStyle = workbookDst.createCellStyle();
        Row rowDst         = null;
        Cell cellDst       = null;

        Sheet sheetSrc = workbookSrc.getSheetAt(fromSheetIndex);
        Sheet sheetDst = workbookDst.createSheet(sheetSrc.getSheetName());

        // process each Row
        for (int rowIndex = 0; rowIndex < sheetSrc.getPhysicalNumberOfRows(); rowIndex++) {
            rowDst = sheetDst.createRow(rowIndex);
            rowDst.setHeight(sheetSrc.getRow(rowIndex).getHeight());

            // process each Col
            for (int colIndex = 0; colIndex < sheetSrc.getRow(rowIndex).getPhysicalNumberOfCells(); colIndex++) {
                Cell cellSrc = sheetSrc.getRow(rowIndex).getCell(colIndex);
                cellDst      = rowDst.createCell(colIndex);

                copyCells(cellSrc, cellDst);

                //Cell cellSrc = sheetSrc.getRow(rowIndex).getCell(colIndex);
                sheetDst.setColumnWidth(colIndex, sheetSrc.getColumnWidth(colIndex));

//                CellStyle origStyle = cellSrc.getCellStyle();

                //newStyle.cloneStyleFrom(origStyle);
//                cellDst.setCellStyle(origStyle);

                cellDst.setCellComment(cellSrc.getCellComment());

//                switch (cellSrc.getCellType()) {
//                    case STRING:
//                        cellDst.setCellValue(cellSrc.getRichStringCellValue().getString());
//                        break;
//                    case NUMERIC:
//                        if (DateUtil.isCellDateFormatted(cellDst)) {
//                            cellDst.setCellValue(cellSrc.getDateCellValue());
//                        } else {
//                            cellDst.setCellValue(cellSrc.getNumericCellValue());
//                        }
//                        break;
//                    case BOOLEAN:
//                        cellDst.setCellValue(cellSrc.getBooleanCellValue());
//                        break;
//                    case FORMULA:
//                        cellDst.setCellFormula(cellSrc.getCellFormula());
//                        break;
//                    case ERROR:
//                        cellDst.setCellValue(cellSrc.getErrorCellValue());
//                        break;
//                    case BLANK:
//                        cellDst.setBlank();
//                        break;
//                    default:
//                        Assert.fail("\n\nError: This Celltype is not supported");
//                }
            }
        }

        workbookSrc.close();
        xlSrc.close();

        workbookDst.write(xlDst);
        workbookDst.close();
        xlDst.close();
    }

    private static void copyCells(Cell cellSrc, Cell cellDst) {
        switch (cellSrc.getCellType()) {
            case STRING:
                String string1 = cellSrc.getStringCellValue();
                cellDst.setCellValue(string1);
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cellSrc)) {
                    Date date1 = cellSrc.getDateCellValue();
                    cellDst.setCellValue(date1);
                } else {
                    double cellValue1 = cellSrc.getNumericCellValue();
                    cellDst.setCellValue(cellValue1);
                }
                break;
            case FORMULA:
                String formula1 = cellSrc.getCellFormula();
                cellDst.setCellFormula(formula1);
                break;

            //case : //TODO: further cell types

        }
        copyStyles(cellSrc, cellDst);
    }

    private static void copyStyles(Cell cell1, Cell cell2) {
        CellStyle style1 = cell1.getCellStyle();
        Map<String, Object> properties = new HashMap<String, Object>();

        // CellUtil.ALIGNMENT - Horizontal
        properties.put(CellUtil.ALIGNMENT , style1.getAlignment());
        // CellUtil.VERTICAL_ALIGNMENT - Horizontal
        properties.put(CellUtil.VERTICAL_ALIGNMENT , style1.getVerticalAlignment().getCode());

        //CellUtil.DATA_FORMAT
        short dataFormat1 = style1.getDataFormat();
        if (BuiltinFormats.getBuiltinFormat(dataFormat1) == null) {
            String formatString1 = style1.getDataFormatString();
            DataFormat format2 = cell2.getSheet().getWorkbook().createDataFormat();
            dataFormat1 = format2.getFormat(formatString1);
        }
        properties.put(CellUtil.DATA_FORMAT, dataFormat1);

        //CellUtil.FILL_PATTERN
        //CellUtil.FILL_FOREGROUND_COLOR
        FillPatternType fillPattern = style1.getFillPattern();
        short fillForegroundColor = style1.getFillForegroundColor(); //gets only indexed colors, no custom HSSF or XSSF colors
        properties.put(CellUtil.FILL_PATTERN, fillPattern);
        properties.put(CellUtil.FILL_FOREGROUND_COLOR, fillForegroundColor);

        //CellUtil.FONT
        Font font1 = cell1.getSheet().getWorkbook().getFontAt(style1.getFontIndexAsInt());
        Font font2 = copyFont(font1, cell2.getSheet().getWorkbook());
        properties.put(CellUtil.FONT, font2.getIndexAsInt());

        //BORDERS
        BorderStyle borderStyle = null;
        short borderColor = -1;
        //CellUtil.BORDER_LEFT
        //CellUtil.LEFT_BORDER_COLOR
        borderStyle = style1.getBorderLeft();
        properties.put(CellUtil.BORDER_LEFT, borderStyle);
        borderColor = style1.getLeftBorderColor();
        properties.put(CellUtil.LEFT_BORDER_COLOR, borderColor);
        //CellUtil.BORDER_RIGHT
        //CellUtil.RIGHT_BORDER_COLOR
        borderStyle = style1.getBorderRight();
        properties.put(CellUtil.BORDER_RIGHT, borderStyle);
        borderColor = style1.getRightBorderColor();
        properties.put(CellUtil.RIGHT_BORDER_COLOR, borderColor);
        //CellUtil.BORDER_TOP
        //CellUtil.TOP_BORDER_COLOR
        borderStyle = style1.getBorderTop();
        properties.put(CellUtil.BORDER_TOP, borderStyle);
        borderColor = style1.getTopBorderColor();
        properties.put(CellUtil.TOP_BORDER_COLOR, borderColor);
        //CellUtil.BORDER_BOTTOM
        //CellUtil.BOTTOM_BORDER_COLOR
        borderStyle = style1.getBorderBottom();
        properties.put(CellUtil.BORDER_BOTTOM, borderStyle);
        borderColor = style1.getBottomBorderColor();
        properties.put(CellUtil.BOTTOM_BORDER_COLOR, borderColor);

        //CellUtil
        CellUtil.setCellStyleProperties(cell2, properties);
    }

    private static Font copyFont(Font font1, Workbook wb2) {
        boolean isBold = font1.getBold();
        short color = font1.getColor();
        short fontHeight = font1.getFontHeight();
        String fontName = font1.getFontName();
        boolean isItalic = font1.getItalic();
        boolean isStrikeout = font1.getStrikeout();
        short typeOffset = font1.getTypeOffset();
        byte underline = font1.getUnderline();

        Font font2 = wb2.findFont(isBold, color, fontHeight, fontName, isItalic, isStrikeout, typeOffset, underline);
        if (font2 == null) {
            font2 = wb2.createFont();
            font2.setBold(isBold);
            font2.setColor(color);
            font2.setFontHeight(fontHeight);
            font2.setFontName(fontName);
            font2.setItalic(isItalic);
            font2.setStrikeout(isStrikeout);
            font2.setTypeOffset(typeOffset);
            font2.setUnderline(underline);
        }

        return font2;
    }

}

