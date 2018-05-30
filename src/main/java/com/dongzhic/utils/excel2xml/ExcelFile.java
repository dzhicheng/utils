package com.dongzhic.utils.excel2xml;

/**
 * Description:
 *
 * @author: dongzhic
 * @date: 2018/5/30 18:46
 */
public class ExcelFile {

    private String tableName;
    private String tableMeaning;
    private String excelPath;

    public ExcelFile() {
    }

    public ExcelFile(String tableName, String tableMeaning, String excelPath) {
        this.tableName = tableName;
        this.tableMeaning = tableMeaning;
        this.excelPath = excelPath;
    }

    public String getTableName() {
        return tableName;
    }

    public void setTableName(String tableName) {
        this.tableName = tableName;
    }

    public String getTableMeaning() {
        return tableMeaning;
    }

    public void setTableMeaning(String tableMeaning) {
        this.tableMeaning = tableMeaning;
    }

    public String getExcelPath() {
        return excelPath;
    }

    public void setExcelPath(String excelPath) {
        this.excelPath = excelPath;
    }
}
