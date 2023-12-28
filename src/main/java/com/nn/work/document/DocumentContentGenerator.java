package com.nn.work.document;

public interface DocumentContentGenerator {


    /**
     * 需求文档，生成 需求内容。
     *
     * 源文件：
     *
     * 4a角色新增  输入4a角色信息
     *           保存4a角色信息
     * 结果例如：
     * 4a角色新增
     * 输入4a角色信息，保存4a角色信息
     *
     * @param sourceFilePath
     * @param requirementIndex
     * @param requirementDescIndex
     * @param targetFilePath
     */
    void generateRequirementContent(String sourceFilePath, int requirementIndex, int  requirementDescIndex, String targetFilePath);



    /**
     * 单元测试文档，生成 测试内容。
     *
     * 例如：
     * 1.新增集市库密码
     * 测试新增集市库密码是否符合预期
     *
     * @param sourceFilePath cosmic文档路径，一般是.xlsx
     * @param  columnIndexToRead 需求/功能描述的列
     * @param targetFilePath 生成的文档路径，一般是txt文档
     */
    void generateUnitTestContent(String sourceFilePath, int columnIndexToRead , String targetFilePath);


    /**
     * 单元测试文档，生成 测试结果表格信息。
     *
     * 结果文档一般是csv文件，内容例如：
     * 测试项	测试用例号	测试特性	用例描述	测试结论
     * 新增集市库密码	SP-01-01	功能测试	测试新增集市库密码是否符合预期	通过
     *
     * @param sourceFilePath cosmic文档路径，一般是.xlsx
     * @param columnIndexToRead 需求/功能描述的列
     * @param targetFilePath 生成的文档路径，一般是csv文档
     */
    void generateUnitTestResultTable(String sourceFilePath, int columnIndexToRead, String targetFilePath);


    /**
     * 联调测试文档，生成测试结果，一般是csv文件，内容例如：
     *
     * 序号	测试功能	测试人员
     * 1	新增集市库密码	张旭
     * 2	删除集市库密码	张旭
     * 3	修改集市库密码	张旭
     * 4	查询集市库密码	张旭
     *
     * @param sourceFilePath cosmic文档路径，一般是.xlsx
     * @param columnIndexToRead
     * @param targetFilePath
     */
    void generateIntegrationTestContent(String sourceFilePath, int columnIndexToRead, String targetFilePath);


    /**
     * 联调测试文档，生成测试结果表格，表格较为复杂，例如
     *
     * 测试编号	4.1
     * 测试子项目名称	新增集市库密码
     * 测试目的	是否调用接口成功/交互成功
     * 预设条件	用户登录深度云页面或者有访问深度云接口的权限
     * 测试过程	序号	测试步骤	输入	检查点	备注
     * 	        1、	进入深度分析云系统，调用服务接口		服务接口是否可用
     * 预期结果	可以正常调用接口服务
     * 测试结果	√ 通过 □ 不通过。
     * 测试人员	张旭
     * 测试日期	2023.3.28
     * 备注
     *
     *
     * @param sourceFilePath cosmic文档路径，一般是.xlsx
     * @param columnIndexToRead
     * @param targetFilePath 格式是.xlsx
     */
    void generateIntegrationTestResultTable(String sourceFilePath, int columnIndexToRead, String targetFilePath);


}
