package com.example;

import java.io.FileOutputStream;
import java.util.List;
import java.util.Map;


import org.apache.poi.xssf.usermodel.*;

public class DBToExcel {

    public static void createExcel(Map<String, List<String>> map, List<String> strArray,String tableName) {
        // 第一步，创建一个webbook，对应一个Excel文件
        XSSFWorkbook wb = new XSSFWorkbook();
        // 第二步，在webbook中添加一个sheet,对应Excel文件中的sheet
        XSSFSheet sheet = wb.createSheet("sheet1");
        sheet.setDefaultColumnWidth(20);// 默认列宽
        // 第三步，在sheet中添加表头第0行,注意老版本poi对Excel的行数列数有限制short
        XSSFRow row = sheet.createRow((int) 0);
        // 第四步，创建单元格，并设置值表头 设置表头居中
//        XSSFCellStyle style = wb.createCellStyle();
        // 创建一个居中格式
//        style.setAlignment(XSSFCellStyle);

        // 添加excel title
        XSSFCell cell = null;
        for (int i = 0; i < strArray.size(); i++) {
            cell = row.createCell((short) i);
            cell.setCellValue(strArray.get(i));
//            cell.setCellStyle(style);
        }

        // 第五步，写入实体数据 实际应用中这些数据从数据库得到,list中字符串的顺序必须和数组strArray中的顺序一致
        int i = 0;
        for (String str : map.keySet()) {
            row = sheet.createRow((int) i + 1);
            List<String> list = map.get(str);

            // 第四步，创建单元格，并设置值
            for (int j = 0; j < strArray.size(); j++) {
                row.createCell((short) j).setCellValue(list.get(j));
            }

            // 第六步，将文件存到指定位置
            try {
                if(tableName.contains(".")) {
                    String str1 = tableName.substring(0, tableName.indexOf("."));
                    String str2 = tableName.substring(str1.length() + 1, tableName.length());
                    tableName = str2;
                }
                FileOutputStream fout = new FileOutputStream("./"+tableName+".xls");
                wb.write(fout);
                fout.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
            i++;
        }
        System.out.println("导出结束");
    }
}
