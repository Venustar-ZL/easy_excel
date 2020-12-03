package com.example.easy_excel.jackson;

import com.alibaba.excel.util.DateUtils;
import com.alibaba.fastjson.JSONObject;
import com.example.easy_excel.bean.VslVoy;
import com.example.easy_excel.config_file_test.config.*;
import com.example.easy_excel.utils.CellColorUtil;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.dataformat.yaml.YAMLFactory;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.FileInputStream;
import java.lang.reflect.Field;
import java.text.DecimalFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @Classname JacksonTest
 * @Description TODO
 * @Date 2020/12/2 18:23
 * @Author by ZhangLei
 */
@Component
@Slf4j
public class JacksonTest {

    private ConfigBean configBean;

    private String type = "";

    public static void main(String[] args) {
        JacksonTest jacksonTest = new JacksonTest();
        File file = new File("C:\\Users\\hujingyi\\Desktop\\OOCL.xls");
        jacksonTest.read(file);
    }

    public void read(File file)  {
        try {
            ObjectMapper mapper = new ObjectMapper(new YAMLFactory());
            String path = this.getClass().getResource("/application-OOCL.yml").getFile();
            configBean = mapper.readValue(new File(path), ConfigBean.class);
            importExcel(file);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private List<List> importExcel(File file) throws Exception {
        Workbook wb = null;
        String fileName = file.getName();
        Iterator<Sheet> sheets = null;
        List<List> returnlist = new ArrayList<List>();
        try {
            if (fileName.endsWith("xls")) {
                wb = new HSSFWorkbook(new FileInputStream(file));
//                sheets = wb.iterator();
                sheets = getAllBotHiddenSheet(wb);
                type = "xls";
            } else if (fileName.endsWith("xlsx")) {
                wb = new XSSFWorkbook(new FileInputStream(file));
                sheets = getAllBotHiddenSheet(wb);
                type = "xlsx";
            }
            if (sheets == null) {
                throw new Exception("excel中不含有sheet工作表");
            }

            int count = 0;
            while (sheets.hasNext()) {
                List<VslVoy> list = new ArrayList<>();
                Sheet sheet = sheets.next();
//                if ("PCN1".equals(sheet.getSheetName())) {
//                    list = getCellValue(sheet);
//                }
                count++;
                log.info(sheet.getSheetName() + "---" +count);
                list = getCellValue(sheet);
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

    private List<VslVoy> getCellValue(Sheet sheet) {

        List<VslVoy> list = new ArrayList<>();
        List<VslVoyAttribute> attributeList = parseUniversal();
        boolean contentFlag = false;
        for (int i = sheet.getFirstRowNum(); i < sheet.getLastRowNum(); i++) {
            VslVoy vslVoy = new VslVoy();
            Row row = sheet.getRow(i);

            if (row != null) {
                int j = row.getFirstCellNum();
                Cell cell = row.getCell(j);
                if (isHead(cell)) {
                    log.info("=====> 表头 <=====");
                    log.info(cell.getStringCellValue());
                    contentFlag = true;
                    continue;
                }

                if (isTail(cell)) {
                    log.info("=====> 表尾 <=====");
                    log.info(cell.getStringCellValue());
                    contentFlag = false;

                }

                if (contentFlag) {
                    // 内容解析
                    if (isTitle(cell)) {
                        i++;
                        continue;
                    }

                    for (int k = 0; k < attributeList.size(); k++) {
                        Cell temp = row.getCell(j++);
                        if (temp == null || temp.getCellTypeEnum() == CellType.BLANK) {
                            continue;
                        }
                        VslVoyAttribute attribute = attributeList.get(k);
                        String cellValue = getCellConvertValue(temp);
                        dynamicSet(vslVoy, attribute.getName(), cellValue);
                        j = j + attribute.getLength() - 1;
                    }
                    if (vslVoy.getVesselName() != null) {
                        System.out.println(JSONObject.toJSONString(vslVoy));
                    }
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
            Head head = configBean.getHead();
            boolean headAccurate = head.getAccurate();
            String headColor = head.getColor();
            String headPattern = head.getPattern();
            boolean colorMatchResult = colorMatch(CellColorUtil.getColorByCell(cell, type), headColor);
            boolean patternMatchResult = patternMatch(cell.getStringCellValue(), headPattern);
            return headAccurate ? colorMatchResult && patternMatchResult : colorMatchResult || patternMatchResult;
        } catch (Exception ignored) {
            return false;
        }
    }

    private boolean isTail(Cell cell) {
        // 判断表尾是否需要精确匹配
        try {
            Tail tail = configBean.getTail();
            boolean tailAccurate = tail.getAccurate();
            String tailColor = tail.getColor();
            String tailPattern = tail.getPattern();
            boolean colorMatchResult = colorMatch(CellColorUtil.getColorByCell(cell, type), tailColor);
            boolean patternMatchResult = patternMatch(cell.getStringCellValue(), tailPattern);
            return tailAccurate ? colorMatchResult && patternMatchResult : colorMatchResult || patternMatchResult;
        } catch (Exception ignored) {
            return false;
        }
    }

    private boolean isTitle(Cell cell) {
        try {
            TitleSign titleSign = configBean.getContent().getTitleSign();
            String titleColor = titleSign.getColor();
            String titlePattern = titleSign.getPattern();
            String cellColor = CellColorUtil.getColorByCell(cell, type);
            boolean colorMatch = colorMatch(cellColor, titleColor);
            boolean patternMatch = patternMatch(cell.getStringCellValue(), titlePattern);
            return colorMatch || patternMatch;
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
        if (StringUtils.isBlank(color)) {
            return false;
        }
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

    /**
     * 解析通用属性配置
     */
    private List<VslVoyAttribute> parseUniversal() {
        List<VslVoyAttribute> attributeList = new ArrayList<>();
        String[] universalList = configBean.getContent().getUniversal().split(",");
        for (String universal : universalList) {
            VslVoyAttribute vslVoyAttribute = new VslVoyAttribute();
            vslVoyAttribute.setName(universal.substring(0, universal.indexOf("(")));
            vslVoyAttribute.setLength(Integer.parseInt(universal.substring(universal.indexOf("(") + 1, universal.indexOf(")"))));
            attributeList.add(vslVoyAttribute);
        }
        return attributeList;
    }

    /**
     * 根据属性名调用set方法
     * @param vslVoy  对象
     * @param propertyName  属性名
     * @param value  要插入的属性值
     */
    private void dynamicSet(VslVoy vslVoy, String propertyName, Object value){
        try {
            Field field  = vslVoy.getClass().getDeclaredField(propertyName);
            field.setAccessible(true);
            field.set(vslVoy, value);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public String getCellConvertValue(Cell cell) {
        String cellValue= "";
        CellType cellType = cell.getCellTypeEnum();
        switch(cellType) {
            case STRING :
                cellValue = cell.getStringCellValue().trim();
                break;
            case NUMERIC:
                if (HSSFDateUtil.isCellDateFormatted(cell)) {  //判断日期类型
                    cellValue = DateUtils.format(cell.getDateCellValue(), "yyyy-MM-dd");
                } else {
                    cellValue = new DecimalFormat("#.######").format(cell.getNumericCellValue());
                }
                break;
            case BOOLEAN:
                cellValue = String.valueOf(cell.getBooleanCellValue());
                break;
            case FORMULA:

                try {
                    if (HSSFDateUtil.isCellDateFormatted(cell)) {  //判断日期类型
                        cellValue = DateUtils.format(cell.getDateCellValue(), "yyyy-MM-dd");
                    } else {
                        cellValue = String.valueOf(cell.getNumericCellValue());
                    }
                } catch (IllegalStateException e) {
                    cellValue = String.valueOf(cell.getRichStringCellValue());
                }
                break;
            default:
                cellValue = "";
                break;
        }
        return cellValue;

    }

    public Iterator<Sheet> getAllBotHiddenSheet(Workbook workbook) {
        List<Sheet> sheetList = new ArrayList<>();
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            if(!workbook.isSheetHidden(i)) {
                sheetList.add(workbook.getSheetAt(i));
            }
        }
        return sheetList.iterator();
    }


}
