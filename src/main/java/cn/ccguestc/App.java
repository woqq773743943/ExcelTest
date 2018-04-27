package cn.ccguestc;

import java.io.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

public class App 
{
    public static void main( String[] args )
    {
        XSSFWorkbook workbook = null; // 读取的文件
        XSSFSheet sheet = null;
        try {
            workbook = new XSSFWorkbook();
            sheet = workbook.createSheet("sheet1");
            sheet.addMergedRegion(new CellRangeAddress(0,3,0,0));
            sheet.addMergedRegion(new CellRangeAddress(0,3,3,3));
            sheet.addMergedRegion(new CellRangeAddress(0,3,4,4));
            XSSFRow row = sheet.createRow(0);
            XSSFCellStyle alignStyle = workbook.createCellStyle();
            alignStyle.setAlignment(HorizontalAlignment.CENTER);
            alignStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            row.createCell(0).setCellValue("工作站");
            row.createCell(1).setCellValue("位置");
            row.createCell(2).setCellValue("序号");
            row.createCell(3).setCellValue("订单号");
            row.createCell(4).setCellValue("成品号/型号");
            row.getCell(0).setCellStyle(alignStyle);
            row.getCell(3).setCellStyle(alignStyle);
            row.getCell(4).setCellStyle(alignStyle);
            for (int i = 1; i <= 3; i++){
                XSSFRow eachRow = sheet.createRow(i);
                eachRow.createCell(1).setCellValue("位置");
                eachRow.createCell(2).setCellValue("序号");
            }
           workbook.write(new FileOutputStream(new File("G:/test1.xlsx")));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
