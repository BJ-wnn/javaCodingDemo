package com.nn.demo.office.excel.impl;

import com.nn.demo.office.excel.ExcelFileSimpleReader;
import lombok.NonNull;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

/**
 * @author NanNan Wang
 * @date 2023/12/19
 */

public class ExcelFileSimpleReaderImpl implements ExcelFileSimpleReader {

    @Override
    public void read(String filePath) {
        try (FileInputStream fileInputStream = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fileInputStream)) {

            // 默认读取第一个sheet
            Sheet sheet = workbook.getSheetAt(0);

            // 遍历每一行
            for (Row row : sheet) {
                // 遍历每一列
                for (Cell cell : row) {
                    // 获取单元格的值
                    String cellValue = getCellValue(cell);
                    System.out.print(cellValue + "\t");
                }
                System.out.println(); // 换行
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    @Override
    public int maxRow(String filePath) {
        int maxRow = -1;
        try (FileInputStream fileInputStream = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fileInputStream)) {

            // 默认读取第一个sheet
            Sheet sheet = workbook.getSheetAt(0);

            // 获取最大行数
            maxRow = sheet.getPhysicalNumberOfRows();

        } catch (IOException e) {
            e.printStackTrace();
        }

        return maxRow;
    }

    @Override
    public int maxColumn(@NonNull String filePath) {
        int maxColumn = -1;
        try (FileInputStream fileInputStream = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fileInputStream)) {

            // 默认读取第一个sheet
            Sheet sheet = workbook.getSheetAt(0);

            // 获取最大列数
            Row row = sheet.getRow(0);
            if (row != null) {
                int lastCellNum = row.getLastCellNum();
                if (lastCellNum > maxColumn) {
                    maxColumn = lastCellNum;
                }
            }


        } catch (IOException e) {
            e.printStackTrace();
        }
        return maxColumn;
    }

    // 获取单元格的值
    private static String getCellValue(Cell cell) {
        if (cell == null) {
            return ""; // 如果单元格为空，返回空字符串
        }

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }
}
