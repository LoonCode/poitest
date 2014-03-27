package com.poitest;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.usermodel.HeaderFooter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.*;


public class PoiTest {
    /**
     * @param args
     * @throws Exception
     */
    public static void main(String[] args) throws Exception {

        String excel6 = "D:\\CreateTest.xlsx";
        String excel7 = "D:\\CreateTest.xls";


        createExcel(excel6);
    }


    static void createExcel(String filePath) throws Exception {

        List<String> hospitalList = new ArrayList();
        hospitalList.add("hospital1");
        hospitalList.add("hospital2");
        hospitalList.add("hospital3");
        hospitalList.add("hospital4");

        List<Map<String, String>> list = new ArrayList<Map<String, String>>();
        Map<String, String> map1 = new HashMap<String, String>();
        map1.put("表1", "123");
        map1.put("表2", "1234");
        map1.put("表3", "1235");

        Map<String, String> map2 = new HashMap<String, String>();
        map2.put("表1", "1234443");
        map2.put("表2", "123444");
        map2.put("表3", "123445");

        list.add(map1);
        list.add(map2);


        FileOutputStream fos = new FileOutputStream(new File(filePath));
        Workbook wb = null;
        Row row = null;
        Cell cell = null;
        if (filePath.endsWith("xls"))
            wb = new HSSFWorkbook();
        else
            wb = new XSSFWorkbook();

        CreationHelper createHelper = wb.getCreationHelper();

        for (String str : hospitalList) {

            Sheet sheet = wb.createSheet(str);
            CellStyle cellStyle = wb.createCellStyle();
            row = sheet.createRow(0);
            row.setHeight((short) 300);


            row.createCell(0).setCellValue("周数");
            row.createCell(1).setCellValue("时间");
//            for () {
                cell = row.createCell(2);
                cell.setCellValue("表名");
                cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
                cell.setCellStyle(cellStyle);
//            }


            row = sheet.createRow(1);
            row.setHeight((short) 300);

            row.createCell(2).setCellValue("时间1");
            row.createCell(3).setCellValue("时间2");


//            row.setHeight((short) 250);
            int rowIndex = 2;
            for (Map map : list) {
                int cellIndex = 2;
                row = sheet.createRow(rowIndex);
                row.setHeight((short) 300);

                row.createCell(cellIndex++).setCellValue((String) map.get("表1"));
                row.createCell(cellIndex++).setCellValue((String) map.get("表2"));
                row.createCell(cellIndex++).setCellValue((String) map.get("表3"));
                rowIndex++;
            }

            sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 0));

            sheet.addMergedRegion(new CellRangeAddress(0, 1, 1, 1));

            sheet.addMergedRegion(new CellRangeAddress(0, 0, 2, 4));

        }

//        // 设置冻结列
//        st1.createFreezePane(3, 2);
//


        wb.write(fos);
        fos.close();
    }

    private static void createCell(Workbook wb, Row row, short column, short halign, short valign) {
        Cell cell = row.createCell(column);
        cell.setCellValue("Align It");
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setAlignment(halign);
        cellStyle.setVerticalAlignment(valign);
        cell.setCellStyle(cellStyle);
    }

    private static void createTitel(Workbook wbn) {
//        Cell cell = row.createCell(column);
//        cell.setCellValue("Align It");
//        CellStyle cellStyle = wb.createCellStyle();
//        cellStyle.setAlignment(halign);
//        cellStyle.setVerticalAlignment(valign);
//        cell.setCellStyle(cellStyle);
    }

}
