package com.nn.work.document.impl;

import com.nn.work.document.DocumentContentGenerator;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class DocumentContentGeneratorImplTest {

    // cosmic 文档路径
    private final String cosmicFilePath = "/Users/nanan/Desktop/test/附件5：需求编号：021深度分析云对接4A高敏感功能开发-COSMIC功能拆分表.xlsx";
    private int columnToReadIndex = 8;

    // 单元测试测试项内容文档
    private final String unitTestContentFilePath = "/Users/nanan/Desktop/test/unitContent.txt";


    DocumentContentGenerator documentContentGenerator = new DocumentContentGeneratorImpl();

    @Test
    void generateUnitTestContent() {
        documentContentGenerator.generateUnitTestContent(cosmicFilePath,columnToReadIndex,unitTestContentFilePath);
    }

    // 单元测试测试 结果表格
    private final String unitTestResultTableFilePath = "/Users/nanan/Desktop/test/unitTestResultTable.csv";
    @Test
    void generateUnitTestResultTable() {
        documentContentGenerator.generateUnitTestResultTable(cosmicFilePath,columnToReadIndex,unitTestResultTableFilePath);
    }

    private final String integrationTestContentFilePath = "/Users/nanan/Desktop/test/integrationTestContent.csv";

    @Test
    void generateIntegrationTestContent() {
        documentContentGenerator.generateIntegrationTestContent(cosmicFilePath,columnToReadIndex,integrationTestContentFilePath);
    }


    private final String  integrationTestResultTableFilePath = "/Users/nanan/Desktop/test/integrationTestResultTable.xlsx";
    @Test
    void generateIntegrationTestResultTable() {
        documentContentGenerator.generateIntegrationTestResultTable(cosmicFilePath,columnToReadIndex,integrationTestResultTableFilePath);
    }
}