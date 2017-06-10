package com.huawei.gdp.service;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;

import java.util.Date;
import java.util.List;



import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import com.huawei.gdp.domain.ExcelEntity;


/**
 *一个通用的将List泛型中数据导出为Excel文档的工具类
 */
@Service
public class ExcelExporter {

    /**
     * 根据ExcelEntity等参数生成Workbook 
     *@param <T> 泛型对象
     *@param entity 实体对象
     *@return Workbook Workbook 
     *@throws Exception 抛异常.
     */
    public static <T> Workbook export2Excel(ExcelEntity<T> entity) throws Exception {
        Workbook workbook = export2Excel(entity.getHeader(), entity.getFooter(), entity.getSheetName(),
                entity.getColumnNames(), entity.getMethodNames(), entity.getEntities());
        return workbook;
    }

    /**
     * 根据给定参数导出Excel文档 
     *@param headerTitle
     *            题头
     *@param footerTitle 脚注
     *@param sheetName sheetName
     *@param columnNames
     *            表头名称
     *@param <T> 泛型对象
     *@param methodNames 方法名
     *@param entities 实体类
     *@return Workbook
     *@throws Exception .
     */
    public static <T> Workbook export2Excel(String headerTitle, String footerTitle, String sheetName,
            String[] columnNames, String[] methodNames, List<T> entities) throws Exception {
        if (methodNames.length != columnNames.length) {
            throw new IllegalArgumentException("methodNames.length should be equal to columnNames.length:"
                    + columnNames.length + " " + methodNames.length);
        }
        Workbook newWorkBook2007 = new XSSFWorkbook();
        Sheet sheet = newWorkBook2007.createSheet(sheetName);

        // 设置题头
        Header header = sheet.getHeader();
        header.setCenter(headerTitle);
        // 设置脚注
        Footer footer = sheet.getFooter();
        footer.setCenter(footerTitle);

        int[] columnWidths = new int[columnNames.length];
        // 创建表头
        createTableHeader(sheet, 1, headerTitle, columnNames, columnWidths);
        // 填充表内容
        createTableContent(sheet, 2, methodNames, columnWidths, entities);

        return newWorkBook2007;

    }

    /**
     * 创建表头 
     * @param sheet sheetname
     * @param index 表头开始的行数
     * @param headerTitle 题头
     * @param columnNames 列名
     * @param columnWidths.
     */
    private static void createTableHeader(Sheet sheet, int index, String headerTitle, String[] columnNames,
            int[] columnWidths) {

        Row headerRow = sheet.createRow(index);

        /* 格式设置 */
        // 设置字体
        Font font = sheet.getWorkbook().createFont();
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);// 粗体显示
        // 设置背景色
        CellStyle style = sheet.getWorkbook().createCellStyle();
        style.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setFont(font);

        for (int i = 0; i < columnNames.length; i++) {
            Cell headerCell = headerRow.createCell(i);
            headerCell.setCellStyle(style);
            headerCell.setCellValue(columnNames[i]);
        }

        for (int i = 0; i < columnNames.length; i++) {
            columnWidths[i] = (columnNames[i].getBytes().length + 2) * 256;
            sheet.setColumnWidth(i, columnWidths[i]);
        }

    }

    /**
     * 创建表格内容 
     * @param sheet 表格
     * @param rowIndexBegin
     *            表内容开始的行数
     * @param methodNames
     *            T对象的方法名
     * @param columnWidths 列宽
     * @param entities 实体
     * @throws Exception .
     */
    private static <T> void createTableContent(Sheet sheet, int rowIndexBegin, String[] methodNames, int[] columnWidths,
            List<T> entities) throws Exception {
        Class<? extends Object> clazz = null;
        if (entities.size() > 0) {
            clazz = entities.get(0).getClass();
        }    
        String content = null;
        for (T t : entities) {
            Row row = sheet.createRow(rowIndexBegin++);
            for (int i = 0; i < methodNames.length; i++) {
                Cell cell = row.createCell(i);
                Method method = clazz.getMethod(methodNames[i]);
                Object object = method.invoke(t);
                object = object == null ? "" : object;               
                if (object.getClass().equals(Date.class)) {
                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                    content = sdf.format((Date) object);
                    cell.setCellValue(content);
                } else if (object.getClass().equals(Double.class)) {
                    content = object.toString() + "%";
                    cell.setCellValue(content);
                } else {
                    content = object.toString();
                    cell.setCellValue(content);
                }
                int columnWidth = (content.getBytes().length + 2) * 256;
                if (columnWidth > columnWidths[i]) {
                    columnWidths[i] = columnWidth;
                    sheet.setColumnWidth(i, columnWidths[i]);
                }

            }
        }
    }



    /**
     *将workbook2007保存文件 
     *@param workbook2007 workbook
     *@param dstFile .
     */
    public static void saveWorkBook2007(Workbook workbook2007, String dstFile) {
        File file = new File(dstFile);
        OutputStream os = null;
        try {
            os = new FileOutputStream(file);
            workbook2007.write(os);
        } catch (IOException ce) {
            ce.printStackTrace();
        } finally {
            if (os != null) {
                try {
                    os.close();
                } catch (IOException ce) {
                    System.out.println(ce);
                }
            }
        }
    }

    /**
     * 生成数组
     *@param lists 泛型集合
     *@param cols 列数组
     *@param methods 方法数组
     *@return byte[]
     *@throws Exception .
     */
    public byte[] exportExcel(List<?> lists, String[] cols, String[] methods) throws Exception {
        String fileName = "";
        ExcelEntity<?> excelEntity = new ExcelEntity<>(fileName, cols, methods, lists);
        Workbook excel = ExcelExporter.export2Excel(excelEntity);
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try {
            excel.write(out);
            excel.close();
            out.close();
        } catch (IOException ae) {
            ae.printStackTrace();
        }
        return out.toByteArray();
    }

    
    /**
     * 测试生成excel工具方法
     * @param columnNames 列名
     * @param methodNames 方法名
     * @param <T> 泛型对象
     * @param entities  实体类 
     * @throws Exception .
     */
    public static <T> void testPoi(String[] columnNames, String[] methodNames, List<T> entities) throws Exception {
        String sheetName = "Test";
        String title = "标题栏";
        String dstFile = "d:/temp/test.xlsx";
        Workbook newWorkBook2007 = new XSSFWorkbook();
        Sheet sheet = newWorkBook2007.createSheet(sheetName);
        int[] columnWidths = new int[columnNames.length];
        // 创建表头
        createTableHeader(sheet, 0, title, columnNames, columnWidths);
        // 填充表内容
        createTableContent(sheet, 1, methodNames, columnWidths, entities);
        // 保存为文件
        saveWorkBook2007(newWorkBook2007, dstFile);
        System.out.println("end");

    }

}
