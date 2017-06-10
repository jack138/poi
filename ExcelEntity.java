package com.huawei.gdp.domain;

import java.util.List;

/**
 * 代表要打印的Excel表格,用于存放要导出为Excel的相关数据 
 * @author zhrb@cec.jmu
 * @param <T> 代表要打印的数据实体，如User等
 */
public class ExcelEntity<T> {
    private String sheetName = "Sheet1";// 默认生成的sheet名称
    private String header = "";// 题头
    private String footer = "";// 脚注
    // 底下是必须具备的属性
    private String fileName;
    private String[] columnNames;// 列名
    private String[] methodNames;// 与列名对应的方法名
    private List<T> entities;// 数据实体
    
    /**
     * 初始化表格
     * @param fileName 文件名
     * @param columnNames 列名
     * @param methodNames 方法名
     * @param entities  实体类.
     */
    public ExcelEntity(String fileName, String[] columnNames, String[] methodNames, List<T> entities) {
        this("sheet1", "", "", fileName, columnNames, methodNames, entities);
    }
    
    /**
     * 初始化表格实体
     * @param sheetName sheetName
     * @param header    页头
     * @param footer    页脚
     * @param fileName  文件名
     * @param columnNames  行名
     * @param methodNames  方法名
     * @param entities   实体.
     */
    public ExcelEntity(String sheetName, String header, String footer, String fileName, String[] columnNames,
            String[] methodNames, List<T> entities) {

        this.sheetName = sheetName;
        this.header = header;
        this.footer = footer;
        this.fileName = fileName;
        this.columnNames = columnNames;
        this.methodNames = methodNames;
        this.entities = entities;
    }
    
    /**
     * getHeader
     * @return header.
     */
    public String getHeader() {
        return header;
    }

    /**
     * 设置 header页头
     * @param header 页头.
     */
    public void setHeader(String header) {
        this.header = header;
    }
    
    /**
     * 获取excel表名
     * @return sheetName.
     */
    public String getSheetName() {
        return sheetName;
    }
    
    /**
     * set sheetname
     * @param sheetName sheetName.
     */
    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }
    
    /**
     * getEntities
     * @return list.
     */
    public List<T> getEntities() {
        return entities;
    }

    /**
     * @param entities 用于导出Excel的实体集合.
     */
    public void setEntities(List<T> entities) {
        this.entities = entities;
    }
    
    /**
     * getFooter
     * @return footer.
     */
    public String getFooter() {
        return footer;
    }
    
    /**
     * setFooter
     * @param footer 页脚.
     */
    public void setFooter(String footer) {
        this.footer = footer;
    }
    
    /**
     * get ColumnNames
     * @return columnName.
     */
    public String[] getColumnNames() {
        return columnNames;
    }

    /**
     * set ColumnNames
     * @param columnNames .
     */
    public void setColumnNames(String[] columnNames) {
        this.columnNames = columnNames;
    }

    /**
     * get fileName
     * @return fileName.
     */
    public String getFileName() {
        return fileName;
    }

    /**
     * set fileName
     * @param fileName .
     */
    public void setFileName(String fileName) {
        this.fileName = fileName;
    }
    
    /**
     * get MethodNames
     * @return methodNames.
     */
    public String[] getMethodNames() {
        return methodNames;
    }
    
    /**
     * set MethodNames
     * @param methodNames .
     */
    public void setMethodNames(String[] methodNames) {
        this.methodNames = methodNames;
    }

}
