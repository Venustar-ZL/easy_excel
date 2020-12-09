package com.example.easy_excel.jackson;

import com.alibaba.excel.util.DateUtils;
import com.alibaba.fastjson.JSONObject;
import com.example.easy_excel.bean.PortOfCall;
import com.example.easy_excel.bean.VslVoy;
import com.example.easy_excel.config_file_test.config.*;
import com.example.easy_excel.utils.CellColorUtil;
import com.example.easy_excel.utils.CellUtil;
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
public class LFTest {

    private ConfigBean configBean;

    private String type = "";

    public static void main(String[] args) {
        LFTest jacksonTest = new LFTest();
        File file = new File("C:\\Users\\hujingyi\\Desktop\\2月船期.xlsx");
        jacksonTest.read(file);
    }

    public void read(File file)  {
        try {
            ObjectMapper mapper = new ObjectMapper(new YAMLFactory());
            String path = this.getClass().getResource("/application-LF.yml").getFile();
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

                if (!sheet.getSheetName().contains("拉菲")) {
                    continue;
                }

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
        VslVoyAttribute specialAttribute = parseSpecial();
        Set<Integer> ignoredColumn = new HashSet<>();
        List<Integer> portOfCallRange = new ArrayList<>();
        List<PortOfCall> portOfCallList = new ArrayList<>();
        Map<Integer, String> mergedCell = new HashMap<>();
        initRange(portOfCallRange);
        boolean contentFlag = false;
        for (int i = sheet.getFirstRowNum(); i < sheet.getLastRowNum(); i++) {
            VslVoy vslVoy = new VslVoy();
            List<PortOfCall> portOfCalls = new ArrayList<>();
            vslVoy.setPortOfCalls(portOfCalls);
            Row row = sheet.getRow(i);

            if (row != null) {
                int j = row.getFirstCellNum();
                Cell cell = row.getCell(j);
                if (isHead(sheet, cell)) {
                    log.info("=====> 表头 <=====");
                    log.info(cell.getStringCellValue());
                    contentFlag = true;
                    continue;
                }

                if (isTail(sheet, cell)) {
                    log.info("=====> 表尾 <=====");
                    log.info(cell.getStringCellValue());
                    contentFlag = false;
                    initRange(portOfCallRange);
                    mergedCell.clear();

                }

                if (contentFlag) {
                    // 内容解析
                    if (isTitle(sheet, cell)) {
                        getIgnoredColumn(ignoredColumn, row);
                        parseTitle(portOfCallRange, specialAttribute, row);
                        getPostOfCallNameInTitle(portOfCallList, portOfCallRange, row);
                        continue;
                    }

                    boolean specialBeginFlag = false;
                    int portOfCallNo = 0;
                    for (int k = 0; k < attributeList.size(); k++) {

                        // 判断此列是否需要忽略
                        if (ignoredColumn.contains(j)) {
                            j++;
                        }

                        String cellValue = "";
                        if (mergedCell.containsKey(j)) {
                            cellValue = mergedCell.get(j);
                            j++;
                        }
                        else {

                            Cell temp = row.getCell(j++);
                            if (temp == null) {
                                continue;
                            }

                            // 判断是否为合并单元格
                            boolean mergedRegionFlag = CellUtil.isMergedRegion(sheet, temp.getRowIndex(), temp.getColumnIndex());
                            if (mergedRegionFlag) {
                                mergedCell.put(j - 1, getCellConvertValue(temp));
                            }

                            cellValue = getCellConvertValue(temp);
                        }

                        VslVoyAttribute universalAttribute = attributeList.get(k);

                        // 判断是否进入特殊属性范围,则需进行特殊属性的处理，处理完成之后，不影响通用属性的处理顺序
                        if (universalAttribute.getName().equals(specialAttribute.getBegin())) {
                            specialBeginFlag = true;
                        }

                        if (universalAttribute.getLength() > 1) {
                            for (int start = j; start < j + universalAttribute.getLength(); start++) {
                                Cell c = row.getCell(start);
                                if (c == null) {
                                    continue;
                                }
                                String value = getCellConvertValue(c);
                                cellValue = cellValue + " " + value;
                            }
                            j++;
                        }

                        dynamicSet(vslVoy, universalAttribute.getName(), cellValue);
                        j = j + universalAttribute.getLength() - 1;

                        if ("null".equals(specialAttribute.getEnd())) {
                            if (specialBeginFlag) {
                                portOfCallNo++;
                                dynamicListAdd(vslVoy, cellValue, portOfCallNo, portOfCallList);
                                k--;
                                j++;
                            }
                        }
                        else {
                            if (j >= portOfCallRange.get(0) + 1 && j <= portOfCallRange.get(1) + 1) {
                                portOfCallNo++;
                                dynamicListAdd(vslVoy, cellValue, portOfCallNo, portOfCallList);
                                k--;
                            }
                        }

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
    private boolean isHead(Sheet sheet, Cell cell) {
        Cell nextCell = getHeadNextRangeCell(sheet, cell);
        // 判断表头是否需要精确匹配
        try {
            Head head = configBean.getHead();
            boolean headAccurate = head.getAccurate();
            String headColor = head.getColor();
            String headPattern = head.getPattern();
            String headStyle = head.getStyle();
            boolean colorMatchResult = colorMatch(CellColorUtil.getColorByCell(cell, type), headColor);
            boolean patternMatchResult = patternMatch(cell.getStringCellValue(), headPattern);
            boolean styleMatchResult = styleMatch(cell, nextCell, headStyle);
            return headAccurate ? colorMatchResult && patternMatchResult && styleMatchResult
                                : colorMatchResult || patternMatchResult || styleMatchResult;
        } catch (Exception ignored) {
            return false;
        }
    }


    /**
     * 表尾判断
     * @param cell
     * @return
     */
    private boolean isTail(Sheet sheet, Cell cell) {
        Cell nextCell = getTailNextRangeCell(sheet, cell);
        // 判断表尾是否需要精确匹配
        try {
            Tail tail = configBean.getTail();
            boolean tailAccurate = tail.getAccurate();
            String tailColor = tail.getColor();
            String tailPattern = tail.getPattern();
            String tailStyle = tail.getStyle();
            boolean colorMatchResult = colorMatch(CellColorUtil.getColorByCell(cell, type), tailColor);
            boolean patternMatchResult = patternMatch(cell.getStringCellValue(), tailPattern);
            boolean styleMatchResult = styleMatch(cell, nextCell, tailStyle);
            return tailAccurate ? colorMatchResult && patternMatchResult && styleMatchResult
                                : colorMatchResult || patternMatchResult || styleMatchResult;
        } catch (Exception ignored) {
            return false;
        }
    }

    /** 标题判断
     * @param cell
     * @return
     */
    private boolean isTitle(Sheet sheet, Cell cell) {
        Cell nextCell = getTailNextRangeCell(sheet, cell);
        try {
            TitleSign titleSign = configBean.getContent().getTitleSign();
            String titleColor = titleSign.getColor();
            String titlePattern = titleSign.getPattern();
            String cellColor = CellColorUtil.getColorByCell(cell, type);
            String style = titleSign.getStyle();
            boolean colorMatch = colorMatch(cellColor, titleColor);
            boolean patternMatch = patternMatch(cell.getStringCellValue(), titlePattern);
            boolean styleMatch = styleMatch(cell, nextCell, style);
            return colorMatch || patternMatch || styleMatch;
        } catch (Exception ignored) {
            return false;
        }
    }

    /**
     * 获取表头当前单元格下一行对应的单元格
     * @param sheet
     * @param cell
     * @return
     */
    private Cell getHeadNextRangeCell(Sheet sheet, Cell cell) {
        try {
            Integer range = configBean.getHead().getRange();
            Row nextRow = sheet.getRow(cell.getRowIndex() + range);
            Cell nextCell = nextRow.getCell(cell.getColumnIndex());
            return nextCell;
        } catch (Exception ignored) {
            return null;
        }
    }

    /**
     * 获取表尾当前单元格下一行对应的单元格
     * @param sheet
     * @param cell
     * @return
     */
    private Cell getTailNextRangeCell(Sheet sheet, Cell cell) {
        try {
            Integer range = configBean.getTail().getRange();
            Row nextRow = sheet.getRow(cell.getRowIndex() - range);
            Cell nextCell = nextRow.getCell(cell.getColumnIndex());
            return nextCell;
        } catch (Exception ignored) {
            return null;
        }
    }


    /**
     * 解析忽略的列
     * @param row
     * @return
     */
    private Set<Integer> getIgnoredColumn(Set<Integer> ignoredColumn, Row row) {
        for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);
            if (cell == null) {
                continue;
            }
            String cellValue = getCellConvertValue(cell);
            String[] patterns = configBean.getContent().getIgnoreColumn().split(",");
            for (String pattern : patterns) {
                Pattern p = Pattern.compile(pattern);
                Matcher matcher = p.matcher(cellValue);
                if (matcher.find()) {
                    ignoredColumn.add(cell.getColumnIndex());
                }
            }
        }
        return ignoredColumn;
    }

    /**
     * 解析标题
     * @param portOfCallRange
     * @param specifiedAttribute
     */
    private void parseTitle(List<Integer> portOfCallRange, VslVoyAttribute specifiedAttribute, Row row) {
        int begin = portOfCallRange.get(0);
        int end = portOfCallRange.get(1);
        for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);
            if (cell == null) {
                continue;
            }
            if (getCellConvertValue(cell).contains(specifiedAttribute.getBegin())) {
                begin = i + 1;
            }
            else if (getCellConvertValue(cell).contains(specifiedAttribute.getEnd())) {
                end = i - 1;
            }
        }
        portOfCallRange.set(0, begin);
        portOfCallRange.set(1, end);
    }

    /**
     * 获取标题中的挂靠港名称
     * @param portOfCallList
     * @param portOfCallRange
     * @param row
     */
    private void getPostOfCallNameInTitle(List<PortOfCall> portOfCallList, List<Integer> portOfCallRange, Row row) {
        for (int i = portOfCallRange.get(0); i <= portOfCallRange.get(1); i++) {
            PortOfCall portOfCall = new PortOfCall();
            Cell cell = row.getCell(i);
            String cellValue = getCellConvertValue(cell);
            portOfCall.setPortOfCallName(cellValue);
            portOfCallList.add(portOfCall);
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
     * 进行单元格的样式匹配
     * @param cell
     * @param style
     * @return
     */
    private Boolean styleMatch(Cell cell, Cell nextCell, String style) {
        if (StringUtils.isBlank(style)) {
            return false;
        }
        String[] styles = style.split(",");
        CellStyle cellStyle = cell.getCellStyle();
        CellStyle nextCellStyle = nextCell.getCellStyle();
        return cellStyle.getBorderBottomEnum().name().equals(styles[0])
                && nextCellStyle.getBorderBottomEnum().name().equals(styles[1]);
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
     * 解析特殊属性配置
     */
    private VslVoyAttribute parseSpecial() {
        String special = configBean.getContent().getSpecial();
        VslVoyAttribute vslVoyAttribute = new VslVoyAttribute();
        vslVoyAttribute.setName(special.substring(0, special.indexOf("(")));
        vslVoyAttribute.setBegin(special.substring(special.indexOf("(") + 1, special.indexOf("-")));
        vslVoyAttribute.setEnd(special.substring(special.indexOf("-") + 1, special.indexOf(")")));
        return vslVoyAttribute;
    }

    /**
     * 根据属性名调用set方法
     * @param vslVoy  对象
     * @param propertyName  属性名
     * @param value  要插入的属性值
     */
    private void dynamicSet(VslVoy vslVoy, String propertyName, Object value){
        if ("null".equals(propertyName)) {
            return;
        }
        try {
            Field field  = vslVoy.getClass().getDeclaredField(propertyName);
            field.setAccessible(true);
            field.set(vslVoy, value);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void dynamicListAdd(VslVoy vslVoy, String value, Integer portOfCallNo, List<PortOfCall> portOfCallList) {
        PortOfCall portOfCall = new PortOfCall();
        portOfCall.setEta(value);
        portOfCall.setPortOfCallNo(String.valueOf(portOfCallNo));
        portOfCall.setPortOfCallName(portOfCallList.get(portOfCallNo - 1).getPortOfCallName());
        vslVoy.getPortOfCalls().add(portOfCall);
    }

    /**
     * 获取转换后的单元格值
     * @param cell
     * @return
     */
    public String getCellConvertValue(Cell cell) {
        String cellValue= "";
        CellType cellType;
        try {
            cellType = cell.getCellTypeEnum();
        } catch (Exception e) {
            cellType = CellType.STRING;
        }
        switch(cellType) {
            case STRING :
                cellValue = cell.getStringCellValue().trim();
                break;
            case NUMERIC:
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
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
                    if (HSSFDateUtil.isCellDateFormatted(cell)) {
                        cellValue = DateUtils.format(cell.getDateCellValue(), "yyyy-MM-dd");
                    } else {
                        cellValue = String.valueOf(cell.getNumericCellValue());
                    }
                } catch (IllegalStateException e) {
                    cell.setCellType(CellType.STRING);
                    cellValue = String.valueOf(cell.getRichStringCellValue());
                }
                break;
            default:
                cellValue = "";
                break;
        }
        return cellValue;

    }

    /**
     * 获取所有不隐藏的sheet
     * @param workbook
     * @return
     */
    public Iterator<Sheet> getAllBotHiddenSheet(Workbook workbook) {
        List<Sheet> sheetList = new ArrayList<>();
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            if(!workbook.isSheetHidden(i)) {
                sheetList.add(workbook.getSheetAt(i));
            }
        }
        return sheetList.iterator();
    }

    private void initRange(List<Integer> range) {
        if (range.size() == 0) {
            range.add(0);
            range.add(0);
        }
        else {
            range.set(0, 0);
            range.set(1, 0);
        }
    }


}
