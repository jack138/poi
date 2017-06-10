package com.huawei.gdp.test;

import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;

import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ImportExcel {
    
    
    public static <T> byte[] loadScoreInfo(List<T> lists, String xlsPath, String[] methodNames) throws Exception{  
      //  List temp = new ArrayList();  
        FileInputStream fileIn = new FileInputStream(xlsPath);  
        //根据指定的文件输入流导入Excel从而产生Workbook对象  
        Workbook wb0 = new XSSFWorkbook(fileIn);  
        //获取Excel文档中的第一个表单  
        Sheet sheet = wb0.getSheetAt(0); 
        
        createTableContent(sheet, 1, methodNames, lists);
                
        //FileOutputStream fileOut = new FileOutputStream("hs.xlsx");
        //wb0.write(fileOut);
        //fileOut.close();
        //wb0.close();
        
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try {
            wb0.write(out);
            wb0.close();
            out.close();
        } catch (IOException ae) {
            ae.printStackTrace();
        }
        return out.toByteArray();
  
    }
    
    public static void main(String[] args) throws IOException, NoSuchMethodException, SecurityException, IllegalAccessException, IllegalArgumentException, InvocationTargetException {
        //String path = "c:/source.xlsx";
        //loadScoreInfo(path);
        //System.out.println("ok");
        Teacher t1 = new Teacher();
        t1.setAge(35);
        t1.setName("zj");
        t1.setSex("男");
        Teacher t2 = new Teacher();
        t2.setAge(31);
        t2.setName("hs");
        t2.setSex("女");
        List<Teacher> ts = new ArrayList<Teacher>();
        ts.add(t1);
        ts.add(t2);
        String[] methodNames = {"getName","getAge","getSex"};
        downloadExcel(ts, "", methodNames);
        
        
        
        
    }
    
    
    public static <T> void downloadExcel(List<T> lists, String xlsPath, String[] methodNames) throws NoSuchMethodException, 
    SecurityException, IllegalAccessException, IllegalArgumentException, InvocationTargetException{
        
        Class<? extends Object> clazz = null;
        if (lists.size() > 0) {
            clazz = lists.get(0).getClass();
        } 
        
        String content = null;
        for (T t : lists) {
            
            for (int i = 0; i < methodNames.length; i++) {
                Method method = clazz.getMethod(methodNames[i]);
                Object object = method.invoke(t);
                object = object == null ? "" : object;   
                content = object.toString();
                System.out.println(methodNames[i] + ":" + content);
            }
            
            
        }
        
    }
    
    
    private static <T> void createTableContent(Sheet sheet, int rowIndexBegin, String[] methodNames, List<T> entities)
                    throws Exception {
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

                /*
                 * int columnWidth = (content.getBytes().length + 2) * 256; if
                 * (columnWidth > columnWidths[i]) { columnWidths[i] =
                 * columnWidth; sheet.setColumnWidth(i, columnWidths[i]); }
                 */
            }
        }
    }

}
