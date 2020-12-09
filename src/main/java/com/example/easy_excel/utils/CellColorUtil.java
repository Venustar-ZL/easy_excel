package com.example.easy_excel.utils;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

/**
 * @Classname CellColorUtil
 * @Description 获取单元格背景色
 * @Date 2020/12/2 13:26
 * @Author by ZhangLei
 */
public class CellColorUtil {

    /**
     * 获取颜色
     *
     * @param cell
     * @return
     */
    public static String getColorByCell(Cell cell, String type) {
        StringBuilder sb = new StringBuilder();
        CellStyle style = cell.getCellStyle();

        if ("xls".equals(type)) {
            short color = style.getFillForegroundColor();
            HSSFWorkbook hb = new HSSFWorkbook();
            HSSFColor hc = hb.getCustomPalette().getColor(color);
            short[] s = hc.getTriplet();
            if (s != null) {
                sb.append(s[0]).append(",").append(s[1]).append(",").append(s[2]);
            }
        } else {


            XSSFColor color = (XSSFColor) style.getFillForegroundColorColor();
            if (color != null) {
                if (color.isRGB()) {
                    byte[] bytes = color.getRGB();
                    if (bytes != null && bytes.length == 3) {
                        for (int i = 0; i < bytes.length; i++) {
                            byte b = bytes[i];
                            int temp;
                            if (b < 0) {
                                temp = 256 + (int) b;
                            } else {
                                temp = b;
                            }
                            sb.append(temp);
                            if (i != bytes.length - 1) {
                                sb.append(",");
                            }
                        }
                    }
                }
            }
        }
        return sb.toString();

    }

    /**
     * 获取字体颜色
     * @param cell
     * @return
     */
    public static String getColorByFont(Cell cell) {
        return String.valueOf(cell.getCellStyle().getFontIndex());
    }
}
