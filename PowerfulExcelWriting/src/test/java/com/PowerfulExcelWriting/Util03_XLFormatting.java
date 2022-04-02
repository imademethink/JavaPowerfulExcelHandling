package com.PowerfulExcelWriting;

import com.itextpdf.text.*;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.junit.Assert;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;

public class Util03_XLFormatting {

    public static int mMergeFollowingRegion(String sExcelFile, String sheetName, int firstRow, int lastRow, int firstCol, int lastCol){
        int nMergedRegionCount = -1;
        try {
            FileInputStream oInStream  = new FileInputStream(new File(sExcelFile));
            XSSFWorkbook oWrkBook      = new XSSFWorkbook(oInStream);
            XSSFSheet oWrkSheet        = oWrkBook.getSheet(sheetName);

            nMergedRegionCount = oWrkSheet.addMergedRegion(new CellRangeAddress(
                    firstRow,  //first row (0-based)
                    lastRow,  //last row (0-based)
                    firstCol,  //first column (0-based)
                    lastCol   //last column (0-based)
            ));
            oInStream.close();
            FileOutputStream oOutStream = new FileOutputStream(new File(sExcelFile));
            oWrkBook.write(oOutStream);
            oWrkBook.close();
            oOutStream.close();
        }catch(IOException eXl){
            Assert.fail("UtilLog: Given excel file parsing error");
        }
        return nMergedRegionCount;
    }

    public static void mUnMergeFollowingRegion(String sExcelFile, String sheetName, int nMergedRegionCount){
        if (-1 == nMergedRegionCount){
            Assert.fail("UtilLog: This merged region count -1 is invalid");
        }
        try {
            FileInputStream oInStream  = new FileInputStream(new File(sExcelFile));
            XSSFWorkbook oWrkBook      = new XSSFWorkbook(oInStream);
            XSSFSheet oWrkSheet        = oWrkBook.getSheet(sheetName);

            oWrkSheet.removeMergedRegion(nMergedRegionCount);

            oInStream.close();
            FileOutputStream oOutStream = new FileOutputStream(new File(sExcelFile));
            oWrkBook.write(oOutStream);
            oWrkBook.close();
            oOutStream.close();
        }catch(IOException eXl){
            Assert.fail("UtilLog: Given excel file parsing error");
        }
    }

    public static int mGetTotalMergedRegion(String sExcelFile, String sheetName){
        int nTotalMergedRegionCount = -1;
        try {
            FileInputStream oInStream  = new FileInputStream(new File(sExcelFile));
            XSSFWorkbook oWrkBook      = new XSSFWorkbook(oInStream);
            XSSFSheet oWrkSheet        = oWrkBook.getSheet(sheetName);

            nTotalMergedRegionCount = oWrkSheet.getNumMergedRegions();

            oInStream.close();
            FileOutputStream oOutStream = new FileOutputStream(new File(sExcelFile));
            oWrkBook.write(oOutStream);
            oWrkBook.close();
            oOutStream.close();
        }catch(IOException eXl){
            Assert.fail("UtilLog: Given excel file parsing error");
        }
        return nTotalMergedRegionCount;
    }

    public static void mGetMergedRegionInfo(String sExcelFile, String sheetName, int nMergedRegionCount){
        if (-1 == nMergedRegionCount){
            Assert.fail("UtilLog: This merged region count -1 is invalid");
        }
        try {
            FileInputStream oInStream  = new FileInputStream(new File(sExcelFile));
            XSSFWorkbook oWrkBook      = new XSSFWorkbook(oInStream);
            XSSFSheet oWrkSheet        = oWrkBook.getSheet(sheetName);

            CellRangeAddress oCRA = oWrkSheet.getMergedRegion(nMergedRegionCount);
            System.out.println("UtilLog: For merged region " + nMergedRegionCount);
            System.out.println("UtilLog: First colmn       " + oCRA.getFirstColumn());
            System.out.println("UtilLog: Last colmn        " + oCRA.getLastColumn());
            System.out.println("UtilLog: First row         " + oCRA.getFirstRow());
            System.out.println("UtilLog: Last row          " + oCRA.getLastRow());

            XSSFCell oRequiredCellValue = oWrkSheet.getRow(oCRA.getFirstRow()).getCell(oCRA.getFirstColumn());

            System.out.println("UtilLog: Cell value received " + oRequiredCellValue );
            System.out.println("UtilLog: Content (String)    " + oRequiredCellValue.getStringCellValue());

            oInStream.close();
            FileOutputStream oOutStream = new FileOutputStream(new File(sExcelFile));
            oWrkBook.write(oOutStream);
            oWrkBook.close();
            oOutStream.close();
        }catch(IOException eXl){
            Assert.fail("UtilLog: Given excel file parsing error");
        }
    }

    public static void mPrintDataRow(HashMap<String, Object> hmDataRow){
        for (String sKey : hmDataRow.keySet()) {
            System.out.println(sKey + " : " + hmDataRow.get(sKey));
        }
    }

    /*
     params :=
        sExcelFile     := valid xlsx file path
        sSheetName     := valid xlsx-sheet name
        nRowNum        := valid row count
        nColNum        := valid column count
     returns :=
        XSSFCell       := generic cell content
     asserts :=
        if xlsx file reading error
        if worksheet does not exist
        if worksheet does not have sufficient rows
        if worksheet does not have sufficient columns
     */

    public static XSSFCell mExcel_GetCellValue(String sExcelFile, String sSheetName, int nRowNum, int nColNum){
        XSSFCell oRequiredCellValue = null;
        try {
            FileInputStream oInStream = new FileInputStream(new File(sExcelFile));
            XSSFWorkbook oWrkBook     = new XSSFWorkbook(oInStream);
            XSSFSheet oWrkSheet       = oWrkBook.getSheet(sSheetName);
            if(null == oWrkSheet){
                Assert.fail("UtilLog: Given excel file given sheet does not exist " + sSheetName);
            }
            if(nRowNum > oWrkSheet.getLastRowNum()){
                Assert.fail("UtilLog: Given excel file actual row count is less than required row count");
            }
            if(nColNum > oWrkSheet.getRow(nRowNum).getLastCellNum()){
                Assert.fail("UtilLog: Given excel file actual col count is less than required col count");
            }
            oRequiredCellValue = oWrkSheet.getRow(nRowNum).getCell(nColNum);
            oWrkBook.close();
            oInStream.close();
        }catch(IOException eXl){
            Assert.fail("UtilLog: Given excel file parsing error");
        }
        return oRequiredCellValue;
    }

    /*
     params :=
        sExcelFile     := valid xlsx file path
        sSheetName     := valid xlsx-sheet name
        nRowNum        := valid row count
        nColNum        := valid column count
        objCellValue   := valid cell value (int, float, double, boolean, blank)
     returns :=
        XSSFCell       := generic cell content
     asserts :=
        if xlsx file reading error
        if worksheet does not exist
        if worksheet does not have sufficient rows
        if worksheet does not have sufficient columns
     */

    public static XSSFCell mExcel_SetCellValue(String sExcelFile, String sSheetName, int nRowNum, int nColNum, Object objCellValue){
        XSSFCell oRequiredCellValue = null;
        try {
            FileInputStream oInStream  = new FileInputStream(new File(sExcelFile));
            XSSFWorkbook oWrkBook      = new XSSFWorkbook(oInStream);
            XSSFSheet oWrkSheet        = oWrkBook.getSheet(sSheetName);
            boolean bFreshWorksheet    = false;
            if(null == oWrkSheet){
                oWrkBook.createSheet(sSheetName);
                oWrkSheet = oWrkBook.getSheet(sSheetName);
                oWrkSheet.createRow(nRowNum);
                oWrkSheet.getRow(nRowNum).createCell(nColNum);
                bFreshWorksheet = true;
            }
            if(nRowNum > oWrkSheet.getLastRowNum() && !bFreshWorksheet){
                oWrkSheet.createRow(nRowNum);
            }
            if(nColNum >= oWrkSheet.getRow(nRowNum).getLastCellNum() && !bFreshWorksheet){
                oWrkSheet.getRow(nRowNum).createCell(nColNum);
            }
            oRequiredCellValue = oWrkSheet.getRow(nRowNum).getCell(nColNum);

            if( (objCellValue instanceof Integer ) || (objCellValue instanceof Long ) ){
                oRequiredCellValue.setCellType(CellType.NUMERIC);
                oRequiredCellValue.setCellValue((int)objCellValue); // int or long
            } else
            if( (objCellValue instanceof Float ) || (objCellValue instanceof Double ) ){
                oRequiredCellValue.setCellType(CellType.NUMERIC);
                oRequiredCellValue.setCellValue((float)objCellValue); // float or double
            } else
            if(    (objCellValue instanceof String) && ( ! objCellValue.toString().isEmpty())   ){
                oRequiredCellValue.setCellType(CellType.STRING);
                oRequiredCellValue.setCellValue(objCellValue.toString());
            } else
            if(    (objCellValue instanceof String) && ( objCellValue.toString().isEmpty())     ){
                oRequiredCellValue.setCellType(CellType.BLANK);
                oRequiredCellValue.setCellValue("");
            } else
            if (objCellValue instanceof Boolean){
                oRequiredCellValue.setCellType(CellType.BOOLEAN);
                oRequiredCellValue.setCellValue((boolean)objCellValue);
            }
            // also need to add for objCellValue instanceof Date

            oInStream.close();
            FileOutputStream oOutStream = new FileOutputStream(new File(sExcelFile));
            oWrkBook.write(oOutStream);
            oWrkBook.close();
            oOutStream.close();
        }catch(IOException eXl){
            eXl.printStackTrace();
            Assert.fail("UtilLog: Given excel file parsing error");
        }
        return oRequiredCellValue;
    }

    /*
     params :=
        sExcelFile     := valid xlsx file path
        sSheetName     := valid xlsx-sheet name
        nRowIndex      := valid row count
     returns :=
        HashMap<String, Object> := key-set is 0th row
                                := value-set is for row numbered nRowIndex
     asserts :=
        if xlsx file reading error
     */

    public static HashMap<String, Object> mExcel_ReadDataRow(String sExcelFile, String sSheetName, int nRowIndex){
        HashMap<String, Object> hmDataRow = new HashMap<String, Object>();
        try {
            FileInputStream oInStream = new FileInputStream(new File(sExcelFile));
            XSSFWorkbook oWrkBook     = new XSSFWorkbook(oInStream);
            XSSFSheet oWrkSheet       = oWrkBook.getSheet(sSheetName);

            for (int k=0; k< oWrkSheet.getRow(nRowIndex).getLastCellNum(); k++){
                hmDataRow.put(
                    oWrkSheet.getRow(0).getCell(k).toString(), // key
                    oWrkSheet.getRow(nRowIndex).getCell(k)              // value
                );
            }
            oWrkBook.close();
            oInStream.close();
        }catch(IOException eXl){
            Assert.fail("UtilLog: Given excel file parsing error");
        }
        return hmDataRow;
    }

    /*
     params :=
        sExcelFile     := valid xlsx file path
        sSheetName     := valid xlsx-sheet name
        sUniqueId      := mainly 1 of the ids in 0th column
     returns :=
        HashMap<String, Object> := key-set is 0th row
                                := value-set is for row where oth column matches sUniqueId
     asserts :=
        if xlsx file reading error
     */

    public static HashMap<String, Object> mExcel_ReadDataRow_UniqueId(String sExcelFile, String sSheetName, String sUniqueId){
        HashMap<String, Object> hmDataRow = new HashMap<String, Object>();
        try {
            FileInputStream oInStream = new FileInputStream(new File(sExcelFile));
            XSSFWorkbook oWrkBook     = new XSSFWorkbook(oInStream);
            XSSFSheet oWrkSheet       = oWrkBook.getSheet(sSheetName);
            int nRowIndex             = -1;

            for (int k=0; k< oWrkSheet.getLastRowNum(); k++){
                if(oWrkSheet.getRow(k).getCell(0).toString().matches(sUniqueId)){
                    nRowIndex = k;
                    break;
                }
            }

            for (int k=0; k< oWrkSheet.getRow(nRowIndex).getLastCellNum(); k++){
                hmDataRow.put(
                        oWrkSheet.getRow(0).getCell(k).toString(), // key
                        oWrkSheet.getRow(nRowIndex).getCell(k)              // value
                );
            }
            oWrkBook.close();
            oInStream.close();
        }catch(IOException eXl){
            Assert.fail("UtilLog: Given excel file parsing error");
        }
        return hmDataRow;
    }


    /*
     params :=
        sExcelFile     := valid xlsx file path
        sSheetName     := valid xlsx-sheet name
        sUniqueId      := mainly 1 of the ids in 0th column
        sSpecificColnName := valid column id
     returns :=
        XSSFCell       := Matching cell value
     asserts :=
        if xlsx file reading error
     */

    public static XSSFCell mExcel_ReadCell_UniqueId_SpecificCol(String sExcelFile, String sSheetName,
                                                                String sUniqueId, String sSpecificColnName){
        XSSFCell oRequiredCell            = null;
        HashMap<String, Object> hmDataRow = new HashMap<String, Object>();
        try {
            FileInputStream oInStream = new FileInputStream(new File(sExcelFile));
            XSSFWorkbook oWrkBook     = new XSSFWorkbook(oInStream);
            XSSFSheet oWrkSheet       = oWrkBook.getSheet(sSheetName);

            int nColIndex             = -1;
            XSSFRow oRowHeader        = oWrkSheet.getRow(0);
            for (int b=0; b < oRowHeader.getLastCellNum(); b++){
                if(oRowHeader.getCell(b).toString().contains(sSpecificColnName)){
                    nColIndex = b;
                    break;
                }
            }

            int nRowIndex             = -1;
            for (int k=1; k< oWrkSheet.getLastRowNum(); k++){
                if(oWrkSheet.getRow(k).getCell(0).toString().matches(sUniqueId)){
                    nRowIndex = k;
                    break;
                }
            }

            oRequiredCell = oWrkSheet.getRow(nRowIndex).getCell(nColIndex);

            oWrkBook.close();
            oInStream.close();
        }catch(IOException eXl){
            Assert.fail("UtilLog: Given excel file parsing error");
        }
        return oRequiredCell;
    }


    /*
     params :=
        sExcelFile     := valid xlsx file path
        sSheetName     := valid xlsx-sheet name
     returns :=
        ArrayList<ArrayList<Object>> := each internal ArrayList is each row content
     asserts :=
        if xlsx file reading error
     */

    public static Object[][] mExcel_ReadData2DAry(String sExcelFile, String sSheetName){

        try {
            FileInputStream oInStream = new FileInputStream(new File(sExcelFile));
            XSSFWorkbook oWrkBook     = new XSSFWorkbook(oInStream);
            XSSFSheet oWrkSheet       = oWrkBook.getSheet(sSheetName);
            int nMaxRow               = oWrkSheet.getLastRowNum();
            int nMaxCol               = oWrkSheet.getRow(0).getLastCellNum();
            Object[][] ary2D          = new Object[nMaxRow][nMaxCol];

            for (int r=0; r < oWrkSheet.getLastRowNum(); r++){
                XSSFRow oRow          = oWrkSheet.getRow(r);
                for (int c=0; c < oRow.getLastCellNum(); c++){
                    ary2D[r][c] = oRow.getCell(c);
                    System.out.print(oRow.getCell(c) + " ");
                }
                System.out.println();
            }
            oWrkBook.close();
            oInStream.close();
            return ary2D;
        }catch(IOException eXl){
            Assert.fail("UtilLog: Given excel file parsing error");
            return null;
        }
    }

    /*
     params :=
        sExcelFile     := valid xlsx file path
        sSheetName     := valid xlsx-sheet name
     returns :=
        ArrayList<ArrayList<Object>> := each internal ArrayList is each row content
     asserts :=
        if xlsx file reading error
     */

    public static ArrayList<ArrayList<Object>> mExcel_ReadData2DAryList(String sExcelFile, String sSheetName){
        ArrayList<ArrayList<Object>> arylst2D = new ArrayList<ArrayList<Object> >();
        try {
            FileInputStream oInStream = new FileInputStream(new File(sExcelFile));
            XSSFWorkbook oWrkBook     = new XSSFWorkbook(oInStream);
            XSSFSheet oWrkSheet       = oWrkBook.getSheet(sSheetName);

            for (int row=0; row < oWrkSheet.getLastRowNum(); row++){
                XSSFRow oRow              = oWrkSheet.getRow(row);
                ArrayList<Object> arylst  = new ArrayList<Object>();
                for (int c=0; c < oRow.getLastCellNum(); c++){
                    arylst.add(c,oRow.getCell(c));
                    System.out.print(oRow.getCell(c) + " ");
                }
                System.out.println();
                arylst2D.add(row, arylst);
            }
            oWrkBook.close();
            oInStream.close();
        }catch(IOException eXl){
            Assert.fail("UtilLog: Given excel file parsing error");
        }
        return arylst2D;
    }

    /*
     params :=
        simple code walk through
     returns :=
        nothing
     asserts :=
        none
     */

    public static void mExcel_CellStyle_Explanation(){
        try {
            String basePath            = System.getProperty("user.dir");
            String sExcelPath          = basePath + "\\inputXL\\file_excel_simple_styling.xlsx";
            String sSheetName          = "emp_data";
            FileInputStream oInStream  = new FileInputStream(new File(sExcelPath));
            XSSFWorkbook oWrkBook      = new XSSFWorkbook(oInStream);
            XSSFSheet oWrkSheet        = oWrkBook.getSheet(sSheetName);

            // set all row height (which are empty)
            oWrkSheet.setDefaultRowHeight((short)105);
            // set particular row height
            XSSFRow oRow   = oWrkSheet.createRow(7);
            oRow.setHeight((short) 577);

            // set all column width (which are empty)
            oWrkSheet.setDefaultColumnWidth((short)125);
            // set particular column width
            oWrkSheet.setColumnWidth(5, (short)3445);


            // individual cell style
            XSSFCell oCell = oWrkSheet.getRow(7).createCell(3);
            oCell.setCellValue("Max1111111111111111111111111111111111");
                // styling related classes
                XSSFCellStyle oStyle = oWrkBook.createCellStyle();
                // font
                    XSSFFont oFont = oWrkBook.createFont();
                    oFont.setBold( true );
                    oFont.setItalic(true);
                    //oFont.setColor((short) 246);
                oStyle.setFont(oFont);
                // wrap test
                oStyle.setWrapText(true);
                //oStyle.setShrinkToFit(true);
                oStyle.setVerticalAlignment(VerticalAlignment.TOP);
                oStyle.setAlignment(HorizontalAlignment.LEFT);
            oCell.setCellStyle(oStyle);

            oInStream.close();
            FileOutputStream oOutStream = new FileOutputStream(new File(sExcelPath));
            oWrkBook.write(oOutStream);
            oWrkBook.close();
            oOutStream.close();
        }catch(Exception eXl){
            Assert.fail("UtilLog: Given excel file parsing error");
        }
    }


    /*
     params :=
        simple code walk through
     returns :=
        nothing
     asserts :=
        none
     */

    public static void mExcel_Shifting_Operation(){
        try {
            String basePath            = System.getProperty("user.dir");
            String sExcelPath          = basePath + "\\inputXL\\file_excel_shifting_operation.xlsx";
            String sSheetName          = "shifting";
            FileInputStream oInStream  = new FileInputStream(new File(sExcelPath));
            XSSFWorkbook oWrkBook      = new XSSFWorkbook(oInStream);
            XSSFSheet oWrkSheet        = oWrkBook.getSheet(sSheetName);

            // row shifting in this apache-poi version has a bug - it deletes the rows in question
            // shifting can be done for 1 row or multiple
            // shifting can be done above or below
            //oWrkSheet.shiftRows(4,6, -1);

            // column shifting - need to be careful as to what to be
            // shifting can be done for 1 col or multiple
            // shifting can be done left or right side
            oWrkSheet.shiftColumns(2,4, -1);

            oInStream.close();
            FileOutputStream oOutStream = new FileOutputStream(new File(sExcelPath));
            oWrkBook.write(oOutStream);
            oWrkBook.close();
            oOutStream.close();
        }catch(Exception eXl){
            Assert.fail("UtilLog: Given excel file parsing error");
        }
    }

    /*
     params :=
        simple code walk through
        First see sqlite db select query output
        Then get the same data into ResultSet (a 2D array)
        Then dump it into xlsx file
     returns :=
        nothing
     asserts :=
        none
     */

    public static void mExcel_DumpResultSet_to_Excel(){
//        DOS commands to check SQlite3 table content
//        e:>     cd E:\TonyStark\ExcelHandling\PowerfulExcelWriting\inputXL\simple_sqlite
//        e:>     sqlite3 Tony.db
//        sqlite> .tables
//        sqlite> select * from Employee;
//
//        EmpId | EmpName | EmpSalary | EmpGender
//        100   | Jon     | 10500     | m
//        200   | Max     | 20500     | m
//        300   | July    | 1500      | f


        String url = "jdbc:sqlite:E://TonyStark//ExcelHandling//PowerfulExcelWriting//inputXL//simple_sqlite//Tony.db";
        Connection conn = null;
        ResultSet rs    = null;
        // Convert ResultSet to 2D Array
        Object[][] arr2D= new Object[4][4];

        try {
            Class.forName("org.sqlite.JDBC");
            conn = DriverManager.getConnection(url);
            String sql = "SELECT * FROM Employee";
            PreparedStatement pstmt = conn.prepareStatement(sql);
            rs                      = pstmt.executeQuery();
            ResultSetMetaData rsmd  = rs.getMetaData();
            for(int n=1; n <= rsmd.getColumnCount(); n++){
                arr2D[0][n-1] = rsmd.getColumnName(n);
                System.out.print(rsmd.getColumnName(n) + "  ");
            }
            System.out.println();
            int n=1;
            while (rs.next()) {
                int c = 0;
                arr2D[n][c] = rs.getInt("EmpId");
                System.out.print(arr2D[n][c] +  "    ");
                c++;
                arr2D[n][c] = rs.getString("EmpName");
                System.out.print(arr2D[n][c] +  "      ");
                c++;
                arr2D[n][c] = rs.getInt("EmpSalary");
                System.out.print(arr2D[n][c] +  "      ");
                c++;
                arr2D[n][c] = rs.getString("EmpGender");
                System.out.println(arr2D[n][c]);
                n++;
            }
            for (int r = 0; r < arr2D.length; r++){
                for (int c = 0; c < arr2D[r].length; c++){
                    System.out.print(arr2D[r][c] + " ");
                }
                System.out.println();
            }
        }catch (SQLException | ClassNotFoundException e) {
            e.printStackTrace();
            System.out.println(e.getMessage());
        }


        // Dump 2D array into Excel
        try{
            String basePath            = System.getProperty("user.dir");
            String sExcelFile          = basePath + "\\inputXL\\file_db_to_excel.xlsx";
            FileInputStream oInStream  = new FileInputStream(new File(sExcelFile));
            XSSFWorkbook oWrkBook      = new XSSFWorkbook(oInStream);
            XSSFSheet oWrkSheet        = oWrkBook.createSheet("db_to_excel");

            for (int r = 0; r < arr2D.length; r++){
                oWrkSheet.createRow(r);
                for (int c = 0; c < arr2D[r].length; c++){
                    XSSFCell oCell = oWrkSheet.getRow(r).createCell(c);
                    Object obj     = arr2D[r][c];
                    if( (obj instanceof Integer ) || (obj instanceof Long ) ){
                        oCell.setCellType(CellType.NUMERIC);
                        oCell.setCellValue((int)obj); // int or long
                    } else
                    if( (obj instanceof String) && ( ! obj.toString().isEmpty())   ){
                        oCell.setCellType(CellType.STRING);
                        oCell.setCellValue(obj.toString());
                    }
                }
            }

            oInStream.close();
            FileOutputStream oOutStream = new FileOutputStream(new File(sExcelFile));
            oWrkBook.write(oOutStream);
            oWrkBook.close();
            oOutStream.close();
        }catch(IOException eXl){
            Assert.fail("UtilLog: Given excel file parsing error");
        }
    }



    /*
     params :=
        simple code walk through
        Read xlsx file content into XSSFWorkbook/ XSSFSheet
        Then use iText pdf classes to dump it into pdf
        Need to be cautious about orientation
     returns :=
        PDF file
     asserts :=
        none
     */
    public static void mExcel_ConvertTestCaseExcelToPdf(){
        try{
            String basePath            = System.getProperty("user.dir");
            String sExcelFile          = basePath + "\\inputXL\\excel_to_pdf\\file_excel_text_cases_pdf_to_excel.xlsx";
            String sPdfFile            = basePath + "\\inputXL\\excel_to_pdf\\Excel2PDF_Output.pdf";
            FileInputStream oInStream  = new FileInputStream(new File(sExcelFile));
            XSSFWorkbook oWrkBook      = new XSSFWorkbook(oInStream);
            XSSFSheet oWrkSheet        = oWrkBook.getSheet("login");
            XSSFRow oXSSFRow           = oWrkSheet.getRow(0);

            Document oDocument         = new Document();
            PdfWriter.getInstance(oDocument, new FileOutputStream(sPdfFile));
            //oDocument.setPageCount(2);
            oDocument.setMargins(0.2f, 0.9f,1.4f,5.6f);

            // vertical orientation
            //oDocument.setPageSize(PageSize.A4);

            // horizontal orientation
            oDocument.setPageSize(PageSize.A4.rotate());

            // if specific page size
            //oDocument.setPageSize(new Rectangle(0.0f,0.0f, 400.0f, 100.0f));

            oDocument.open();

            PdfPTable oPdfPTable       = new PdfPTable(oXSSFRow.getLastCellNum());
            // since there are 7 columns in pdf (sExcelFile), so below array of 7 length
            float[] aryColWidth        = {5.0f, 16.0f, 7.0f, 26.0f, 12.0f, 5.0f, 5.0f};
            oPdfPTable.setTotalWidth(aryColWidth);
            oPdfPTable.setTotalWidth(400.0f);
            PdfPCell oPdfPCell         = null;

            Iterator<Row> rowIterator  = oWrkSheet.rowIterator();
            while(rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                while(cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();

                    switch(cell.getCellType()) {
                        case NUMERIC:
                            oPdfPCell=new PdfPCell(new Phrase( Double.toString(cell.getNumericCellValue()) ));
                            break;
                        case STRING:
                            oPdfPCell=new PdfPCell(new Phrase(cell.getStringCellValue()));
                            break;
                        case BOOLEAN:
                            oPdfPCell=new PdfPCell(new Phrase( Boolean.toString(cell.getBooleanCellValue()) ));
                            break;
                    }
                    // for merged cell
                    //oPdfPCell.setColspan(2);
                    //oPdfPCell.setRowspan(2);
                    oPdfPTable.addCell(oPdfPCell);
                    oPdfPCell.setHorizontalAlignment(Element.ALIGN_LEFT);
                }
                oPdfPTable.completeRow();
            }
            oDocument.add(oPdfPTable);
            oDocument.close();
            oWrkBook.close();
            oInStream.close();
        }catch (IOException | DocumentException exPDF) {
            Assert.fail("UtilLog: Given excel file parsing error");
        }

    }














}

