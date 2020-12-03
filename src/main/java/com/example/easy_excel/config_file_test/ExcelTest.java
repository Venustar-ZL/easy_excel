package com.example.easy_excel.config_file_test;

import com.example.easy_excel.config_file_test.config.CNCConfigBean;
import com.example.easy_excel.utils.CellColorUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.FileInputStream;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @Classname ExcelTest
 * @Description TODO
 * @Date 2020/12/2 13:32
 * @Author by ZhangLei
 */
@Component
@Slf4j
public class ExcelTest {

    @Autowired
    private CNCConfigBean CNCConfigBean;

    public List<List> importExcel(File file) throws Exception {
        Workbook wb = null;
        String fileName = file.getName();
        Iterator<Sheet> sheets = null;
        List<List> returnlist = new ArrayList<List>();
        String type = "";
        try {
            if (fileName.endsWith("xls")) {
                wb = new HSSFWorkbook(new FileInputStream(file));
                sheets = wb.iterator();
                type = "xls";
            } else if (fileName.endsWith("xlsx")) {
                wb = new XSSFWorkbook(new FileInputStream(file));
                sheets = wb.iterator();
                type = "xlsx";
            }
            if (sheets == null) {
                throw new Exception("excel中不含有sheet工作表");
            }
            // 遍历excel里每个sheet的数据。
            while (sheets.hasNext()) {
                System.out.println("-----遍历sheet-----");
                Sheet sheet = sheets.next();
                List<Map> list = getCellValue(sheet, type);
                System.out.println(list);
                returnlist.add(list);
            }
        } catch (Exception ex) {
            throw ex;
        } finally {
            if (wb != null) {
                wb.close();
            }
        }
        return returnlist;
    }

    // 获取每一个Sheet工作表中的数。
    private List<Map> getCellValue(Sheet sheet, String type) {
        List<Map> list = new ArrayList<Map>();
        // sheet.getPhysicalNumberOfRows():获取的是物理行数，也就是不包括那些空行（隔行）的情况
        for (int i = sheet.getFirstRowNum(); i < sheet.getLastRowNum(); i++) {
            Map map = new HashMap<>();
            // 获得第i行对象
            Row row = sheet.getRow(i);

            if (row != null) {
                int j = row.getFirstCellNum();

                Cell cell = row.getCell(j);
                if (isHead(cell)) {
                    log.info("=====> 表头 <=====");
                    log.info(cell.getStringCellValue());
                } else if (isTail(cell)) {
                    log.info("=====> 表尾 <=====");
                    log.info(cell.getStringCellValue());
                } else {
                }

            }
        }
        return list;
    }

    /**
     * 表头判断
     *
     * @param cell
     * @return
     */
    private boolean isHead(Cell cell) {
        // 判断表头是否需要精确匹配
        try {
            boolean headAccurate = CNCConfigBean.getHeadAccurate();
            String headColor = CNCConfigBean.getHeadColor();
            String headPattern = CNCConfigBean.getHeadPattern();
            boolean colorMatchResult = colorMatch(CellColorUtil.getColorByCell(cell, "xls"), headColor);
            boolean patternMatchResult = patternMatch(cell.getStringCellValue(), headPattern);
            return headAccurate ? colorMatchResult && patternMatchResult : colorMatchResult || patternMatchResult;
        } catch (Exception ignored) {
            return false;
        }
    }

    private boolean isTail(Cell cell) {
        // 判断表尾是否需要精确匹配
        try {
            boolean tailAccurate = CNCConfigBean.getTailAccurate();
            String tailColor = CNCConfigBean.getTailColor();
            String tailPattern = CNCConfigBean.getTailPattern();
            boolean colorMatchResult = colorMatch(CellColorUtil.getColorByCell(cell, ""), tailColor);
            boolean patternMatchResult = patternMatch(cell.getStringCellValue(), tailPattern);
            return tailAccurate ? colorMatchResult && patternMatchResult : colorMatchResult || patternMatchResult;
        } catch (Exception ignored) {
            return false;
        }
    }

    /**
     * 颜色匹配
     *
     * @param cellColor
     * @param color
     * @return
     */
    private Boolean colorMatch(String cellColor, String color) {
        return StringUtils.equals(cellColor, color);
    }

    /**
     * 进行单元格值与定义的正则表达式的匹配
     *
     * @param cellValue
     * @param pattern
     * @return
     */
    private Boolean patternMatch(String cellValue, String pattern) {
        if (StringUtils.isBlank(pattern)) {
            return false;
        }
        Pattern p = Pattern.compile(pattern);
        Matcher matcher = p.matcher(cellValue);
        return matcher.find();
    }

}
