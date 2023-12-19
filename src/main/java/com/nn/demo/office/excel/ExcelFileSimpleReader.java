package com.nn.demo.office.excel;

import lombok.NonNull;

public interface ExcelFileSimpleReader {

    /**
     * 读取excel文件的第一个sheet的全部内容，并打印到console。
     * @param filePath
     */
    void read(@NonNull String filePath);

    /**
     * 读取excel文件的第一个 sheet 内容的最大行数
     * @param filePath
     * @return
     */
    int maxRow(@NonNull String filePath);

    /**
     * 读取excel文件的第一个 sheet 列数（以第一行为准）
     * @param filePath
     * @return
     */
    int maxColumn(@NonNull String filePath);


}
