package com.nn.demo.office.excel.impl;

import com.nn.demo.office.excel.ExcelFileSimpleReader;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.File;

import static org.junit.jupiter.api.Assertions.*;

class ExcelFileSimpleReaderImplTest {

    private ExcelFileSimpleReader simpleReader = new ExcelFileSimpleReaderImpl();

    private String simpleFilePath;

    @BeforeEach
    public void setUp() {
        // 获取 resources 目录下的文件相对路径
        String relativePath = "src/main/resources/test/simpleExcel.xlsx";

        // 获取项目的工作目录
        String projectDirectory = System.getProperty("user.dir");

        // 拼接得到文件的绝对路径
        String absolutePath = projectDirectory + File.separator + relativePath;

        // 打印文件路径，用于测试
        System.out.println("Absolute Path: " + absolutePath);

        // 在测试方法执行之前进行初始化
        simpleFilePath = absolutePath;
    }

    @Test
    void read() {
        simpleReader.read(simpleFilePath);
    }

    @Test
    void getMaxRow() {
        int maxRow = simpleReader.maxRow(simpleFilePath);
        assertEquals(9,maxRow);
    }

    @Test
    void maxColumn() {
        int maxColumn = simpleReader.maxColumn(simpleFilePath);
        assertEquals(5,maxColumn);
    }
}