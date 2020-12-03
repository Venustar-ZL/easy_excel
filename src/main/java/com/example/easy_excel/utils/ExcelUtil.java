//package com.example.easy_excel.utils;
//
//import java.io.File;
//import java.io.FileInputStream;
//import java.util.*;
//
//import org.apache.commons.lang.StringUtils;
//import org.apache.poi.hssf.usermodel.HSSFWorkbook;
//import org.apache.poi.ss.usermodel.*;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//
//public class ExcelUtil {
//
//
//    public static List<List> importExcel(File file) throws Exception {
//        Workbook wb = null;
//        String fileName = file.getName();// 读取上传文件(excel)的名字，含后缀后
//        // 根据文件后缀名不同(xls和xlsx)获得不同的Workbook实现类对象
//        Iterator<Sheet> sheets = null;
//        List<List> returnlist = new ArrayList<List>();
//        try {
//            if (fileName.endsWith("xls")) {
//                wb = new HSSFWorkbook(new FileInputStream(file));
//                sheets = wb.iterator();
//            } else if (fileName.endsWith("xlsx")) {
//                wb = new XSSFWorkbook(new FileInputStream(file));
//                sheets = wb.iterator();
//            }
//            if (sheets == null) {
//                throw new Exception("excel中不含有sheet工作表");
//            }
//            // 遍历excel里每个sheet的数据。
//            while (sheets.hasNext()) {
//                System.out.println("-----遍历sheet-----");
//                Sheet sheet = sheets.next();
//                List<Map> list = getCellValue(sheet);
//                System.out.println(list);
//                returnlist.add(list);
//            }
//        } catch (Exception ex) {
//            throw ex;
//        } finally {
//            if (wb != null) {
//                wb.close();
//            }
//        }
//        return returnlist;
//    }
//
//
//    // 获取每一个Sheet工作表中的数。
//    private static List<Map> getCellValue(Sheet sheet) {
//        List<Map> list = new ArrayList<Map>();
//        // sheet.getPhysicalNumberOfRows():获取的是物理行数，也就是不包括那些空行（隔行）的情况
//        for (int i = sheet.getFirstRowNum(); i < sheet.getLastRowNum(); i++) {
//            Map map = new HashMap<>();
//            // 获得第i行对象
//            Row row = sheet.getRow(i);
//
//            if (row == null) {
//            } else {
//                int j = row.getFirstCellNum();
//
//                try {
//                    System.out.println(CellColorUtil.getColorByCell(row.getCell(j)));
//                    String rowValue = row.getCell(j).getStringCellValue();
//                    if (StringUtils.isEmpty(rowValue) || rowValue.contains("学生信息")) {
//                        i++;
//                        continue;
//                    }
//                } catch (Exception ignored) {
//
//                }
//
//                Integer id = (int)row.getCell(j++).getNumericCellValue();
//                if (id == null) {
//                    continue;
//                }
//                map.put("学号", id);
//
//                String name = row.getCell(j++).getStringCellValue();
//                if (StringUtils.isEmpty(name)) {
//                    continue;
//                }
//                map.put("姓名", name);
//
//                String sex = row.getCell(j++).getStringCellValue();
//                if (StringUtils.isEmpty(name)) {
//                    continue;
//                }
//                map.put("性别", sex);
//
//                Integer age = (int)row.getCell(j++).getNumericCellValue();
//                if (age == null) {
//                    continue;
//                }
//                map.put("年龄", age);
//
//                list.add(map);
//            }
//        }
//        return list;
//    }
//
//
//
//}