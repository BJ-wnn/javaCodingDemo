package com.nn.work.document.impl;

import com.nn.work.document.DocumentContentGenerator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.StringJoiner;

/**
 * @author NanNan Wang
 * @date 2023/12/20
 */
public class DocumentContentGeneratorImpl implements DocumentContentGenerator {

    @Override
    public void generateRequirementContent(String sourceFilePath, int requirementIndex, int requirementDescIndex, String targetFilePath) {
        try (FileInputStream fis = new FileInputStream(sourceFilePath);
             Workbook workbook = new XSSFWorkbook(fis);) {
            Sheet sheet = workbook.getSheetAt(0);

            // 设置从第2行开始读取数据（0-based index）
            int startRow = 1;

            StringBuilder stringBuilder = new StringBuilder();

            for (int rowIndex = startRow; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if(row == null) {
                    break;
                }
                Cell cell = row.getCell(requirementIndex);
                // 判断 cell 是否在合并区域内
                if (isMergedCell(sheet, cell)) {
                    // 获取合并区域的信息
                    CellRangeAddress mergedRegion = getMergedRegion(sheet, cell);

                    // 输出合并区域的开始行和结束行
                    int startRowMerged = mergedRegion.getFirstRow();
                    int endRowMerged = mergedRegion.getLastRow();
//System.out.println(cell.getStringCellValue());
                    stringBuilder.append(cell.getStringCellValue() +"\n");
                    // 继续读取第9列信息
                    final String descStr = readColumn9Data(sheet, requirementDescIndex, startRowMerged, endRowMerged);
                    stringBuilder.append(descStr + "\n");

                    // 设置下一次迭代的起始行为合并区域的结束行
                    rowIndex = endRowMerged;
                }
                Files.write(Paths.get(targetFilePath), stringBuilder.toString().getBytes());

            }

        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }

    private static String readColumn9Data(Sheet sheet, int columnIndexToRead, int startRow, int endRow) {
//        int columnIndexToRead = 9;  // 第9列，0-based index

        StringJoiner stringJoiner = new StringJoiner("，");
        for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            Cell cell = row.getCell(columnIndexToRead);

            // 输出第9列的值
            String column9Value = (cell != null) ? cell.getStringCellValue() : "";
//            System.out.println("第9列的值: " + column9Value);
            stringJoiner.add(column9Value);
        }
        return stringJoiner.toString();
    }

    private static boolean isMergedCell(Sheet sheet, Cell cell) {
        int rowIndex = cell.getRowIndex();      //该方法用于获取单元格所在的行的索引（行号），索引是从0开始的。
        int columnIndex = cell.getColumnIndex();    //该方法用于获取单元格所在的列的索引（列号），索引同样是从0开始的。

        for (CellRangeAddress mergedRegion : sheet.getMergedRegions()) {
            if (mergedRegion.isInRange(rowIndex, columnIndex)) {
                return true;
            }
        }
        return false;
    }

    private static CellRangeAddress getMergedRegion(Sheet sheet, Cell cell) {
        int rowIndex = cell.getRowIndex();
        int columnIndex = cell.getColumnIndex();

        for (CellRangeAddress mergedRegion : sheet.getMergedRegions()) {
            if (mergedRegion.isInRange(rowIndex, columnIndex)) {
                return mergedRegion;
            }
        }
        return null;
    }

    private static int getColumnIndex(String columnName) {
        // 将列名转换为列索引，比如将 "A" 转换为 0，"B" 转换为 1，以此类推
        return CellReference.convertColStringToIndex(columnName);
    }

    @Override
    public void generateUnitTestContent(String sourceFilePath, int columnIndexToRead, String targetFilePath) {
        try (FileInputStream fileInputStream = new FileInputStream(new File(sourceFilePath));
             Workbook workbook = new XSSFWorkbook(fileInputStream);
             FileWriter writer = new FileWriter(targetFilePath)) {

            // 获取第一个Sheet
            Sheet sheet = workbook.getSheetAt(0);

            int i = 1;

            // 遍历每一行
            for (Row row : sheet) {
                // 获取指定列的单元格
                Cell cell = row.getCell(columnIndexToRead);
                if( null == cell) {
                    break;
                }
                if( cell.getRowIndex() == 0) {
                    continue;
                }
                // 检查单元格是否为空
                if (cell != null && cell.getStringCellValue() != "") {
                    // 读取单元格内容并输出
                    String content = cell.toString();
                    writer.write(i++ + "."+content + "\n");
                    writer.write("测试" + content +"是否符合预期\n");
                }
            }
System.out.println("已生成：" + targetFilePath);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Override
    public void generateUnitTestResultTable(String sourceFilePath, int columnIndexToRead, String targetFilePath) {
        try (FileInputStream fileInputStream = new FileInputStream(new File(sourceFilePath));
             Workbook workbook = new XSSFWorkbook(fileInputStream);
             FileWriter writer = new FileWriter(targetFilePath)) {

            writer.write("测试项,测试用例号,测试特性,用例描述,测试结论\n");

            // 获取第一个Sheet
            Sheet sheet = workbook.getSheetAt(0);

            int i = 1;

            // 遍历每一行
            for (Row row : sheet) {
                // 获取指定列的单元格
                Cell cell = row.getCell(columnIndexToRead);
                if( null == cell) {
                    break;
                }
                if( cell.getRowIndex() == 0) {
                    continue;
                }
                // 检查单元格是否为空
                if (cell != null && cell.getStringCellValue() != "") {
                    // 读取单元格内容并输出
                    String content = cell.toString();
System.out.println(content);
                    StringJoiner stringJoiner = new StringJoiner(",");
                    stringJoiner.add(content);
                    stringJoiner.add("SP-" + String.format("%02d",i++)+"-01");
                    stringJoiner.add("功能测试");
                    stringJoiner.add("测试" + content + "是否符合预期");
                    stringJoiner.add("通过");
                    writer.write(stringJoiner.toString());
                    writer.write("\n");
                }
            }
System.out.println("已生成：" + targetFilePath);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Override
    public void generateIntegrationTestContent(String sourceFilePath, int columnIndexToRead, String targetFilePath) {
        try (FileInputStream fileInputStream = new FileInputStream(new File(sourceFilePath));
             Workbook workbook = new XSSFWorkbook(fileInputStream);
             FileWriter writer = new FileWriter(targetFilePath)) {

            writer.write("序号,测试功能,测试人员\n");

            // 获取第一个Sheet
            Sheet sheet = workbook.getSheetAt(0);

            int i = 1;

            // 遍历每一行
            for (Row row : sheet) {
                // 获取指定列的单元格
                Cell cell = row.getCell(columnIndexToRead);
                if( null == cell) {
                    break;
                }
                if( cell.getRowIndex() == 0) {
                    continue;
                }
                // 检查单元格是否为空
                if (cell != null && cell.getStringCellValue() != "") {
                    // 读取单元格内容并输出
                    String content = cell.toString();
System.out.println(content);
                    StringJoiner stringJoiner = new StringJoiner(",");
                    stringJoiner.add(""+ (i++));
                    stringJoiner.add(content);
                    stringJoiner.add("张旭");
                    writer.write(stringJoiner.toString());
                    writer.write("\n");
                }
            }
System.out.println("已生成：" + targetFilePath);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Override
    public void generateIntegrationTestResultTable(String sourceFilePath, int columnIndexToRead, String targetFilePath) {

        try (FileInputStream fileInputStream = new FileInputStream(new File(sourceFilePath));
             Workbook inputWorkbook = new XSSFWorkbook(fileInputStream);
             FileOutputStream fileOut = new FileOutputStream(targetFilePath);
             Workbook outWorkbook = new XSSFWorkbook()) {
            List<String> contentList = new ArrayList<>(100);
            // 获取第一个Sheet
            Sheet sourceSheet = inputWorkbook.getSheetAt(0);

            // 遍历每一行
            for (Row inputRow : sourceSheet) {
                // 获取指定列的单元格
                Cell inputCell = inputRow.getCell(columnIndexToRead);
                if( null == inputCell) {
                    break;
                }
                if( inputCell.getRowIndex() == 0) {
                    continue;
                }
                // 检查单元格是否为空
                if (inputCell != null && inputCell.getStringCellValue() != "") {
                    // 读取单元格内容并输出
                    String content = inputCell.toString();
                    contentList.add(content);
                }
            }

            // 开始生成对应的表格，一个sheet一个表格
            for(int i = 0 ; i < contentList.size(); i++) {
                final Sheet sheet = outWorkbook.createSheet("Report" + (i + 1));
                // 创建测试报告内容
                createTestReport(sheet,i+1,contentList.get(i));
            }
            // 输出到文件

            outWorkbook.write(fileOut);
System.out.println("测试报告已生成：" + targetFilePath);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    /**
     * 联调测试表格格式
     *
     * @param sheet
     */
    private static void createTestReport(Sheet sheet, int index, String testContent) {

        // 创建标题行
        Row titleRow = sheet.createRow(0);
        titleRow.createCell(0).setCellValue("测试编号");
        titleRow.createCell(1).setCellValue("4."+ index);


        CellRangeAddress mergedRegion = new CellRangeAddress(0, 0, 1, 5);
        sheet.addMergedRegion(mergedRegion);

        Row subtitleRow = sheet.createRow(1);
        subtitleRow.createCell(0).setCellValue("测试子项目名称");
        subtitleRow.createCell(1).setCellValue(testContent);

        CellRangeAddress mergedRegion2 = new CellRangeAddress(1, 1, 1, 5);
        sheet.addMergedRegion(mergedRegion2);

        Row purposeRow = sheet.createRow(2);
        purposeRow.createCell(0).setCellValue("测试目的");
        purposeRow.createCell(1).setCellValue("是否调用接口成功/交互成功");
        CellRangeAddress mergedRegion3 = new CellRangeAddress(2, 2, 1, 5);
        sheet.addMergedRegion(mergedRegion3);

        Row preconditionRow = sheet.createRow(3);
        preconditionRow.createCell(0).setCellValue("预设条件");
        preconditionRow.createCell(1).setCellValue("用户登录深度云页面或者有访问深度云接口的权限");
        CellRangeAddress mergedRegion4 = new CellRangeAddress(3, 3, 1, 5);
        sheet.addMergedRegion(mergedRegion4);

        // 创建测试过程表格
//        Row processHeaderRow = sheet.createRow(5);
        Row processHeaderRow = sheet.createRow(4);
        processHeaderRow.createCell(0).setCellValue("测试过程");
        processHeaderRow.createCell(1).setCellValue("序号");
        processHeaderRow.createCell(2).setCellValue("测试步骤");
        processHeaderRow.createCell(3).setCellValue("输入");
        processHeaderRow.createCell(4).setCellValue("检查点");
        processHeaderRow.createCell(5).setCellValue("备注");

        // 创建测试过程内容
        Row processRow = sheet.createRow(5);
        processRow.createCell(1).setCellValue("1、");
        processRow.createCell(2).setCellValue("进入深度分析云系统，调用服务接口");
        processRow.createCell(4).setCellValue("服务接口是否可用");

        CellRangeAddress mergedRegion5 = new CellRangeAddress(4, 5, 0, 0);
        sheet.addMergedRegion(mergedRegion5);

        // 创建预期结果行
        Row expectedResultRow = sheet.createRow(6);
        expectedResultRow.createCell(0).setCellValue("预期结果");
        expectedResultRow.createCell(1).setCellValue("可以正常调用接口服务");

        CellRangeAddress mergedRegion6 = new CellRangeAddress(6, 6, 1, 5);
        sheet.addMergedRegion(mergedRegion6);

        // 创建测试结果行
        Row testResultRow = sheet.createRow(7);
        testResultRow.createCell(0).setCellValue("测试结果");
        testResultRow.createCell(1).setCellValue("√ 通过 □ 不通过。");

        CellRangeAddress mergedRegion7 = new CellRangeAddress(7, 7, 1, 5);
        sheet.addMergedRegion(mergedRegion7);

        // 创建测试人员行
        Row testerRow = sheet.createRow(8);
        testerRow.createCell(0).setCellValue("测试人员");
        testerRow.createCell(1).setCellValue("张旭");
        CellRangeAddress mergedRegion8 = new CellRangeAddress(8, 8, 1, 5);
        sheet.addMergedRegion(mergedRegion8);

        // 创建测试日期行
        Row testDateRow = sheet.createRow(9);
        testDateRow.createCell(0).setCellValue("测试日期");
        testDateRow.createCell(1).setCellValue("2023.3.28");
        CellRangeAddress mergedRegion9 = new CellRangeAddress(9, 9, 1, 5);
        sheet.addMergedRegion(mergedRegion9);

        // 创建备注行
        Row remarksRow = sheet.createRow(10);
        remarksRow.createCell(0).setCellValue("备注");
        CellRangeAddress mergedRegion10 = new CellRangeAddress(10, 10, 1, 5);
        sheet.addMergedRegion(mergedRegion10);
    }
}
