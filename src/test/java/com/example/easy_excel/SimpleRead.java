package com.example.easy_excel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.example.easy_excel.config_file_test.ExcelTest;
import com.example.easy_excel.jackson.JacksonTest;
import com.example.easy_excel.listener.DemoDataListener;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.File;
import java.util.List;

//import static com.example.easy_excel.utils.ExcelUtil.importExcel;

/**
 * @ClassName:
 * @Description:
 * @Date : 2020-11-30 20:20
 * @Author: ZhangLei
 * @Version: 1.0
 **/
@SpringBootTest
public class SimpleRead {

    @Autowired
    private ExcelTest excelTest;

    @Autowired
    private JacksonTest jacksonTest;

    /**
     * 使用EasyExcel
     */
    @Test
    public void test1() {
        String fileName = "C:\\Users\\hujingyi\\Desktop\\test.xlsx";
        EasyExcel.read(fileName, com.example.easy_excel.bean.DemoData.class, new DemoDataListener()).sheet().doRead();
    }

    /**
     * 使用POI
     */
//    @Test
//    public void test2() {
//        File file = new File("C:\\Users\\hujingyi\\Desktop\\test.xlsx");
//        // File file = new File("E:/2.xls");
//        try {
//            List<List> list = importExcel(file);
//            System.out.println();
//            System.out.println("---------------------分隔---------------------");
//            System.out.println();
//            System.out.println(list);
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
//    }

    /**
     * 自定义配置文件
     */
    @Test
    public void test3() throws Exception {
        File file = new File("C:\\Users\\hujingyi\\Desktop\\test.xlsx");
        List<List> list = excelTest.importExcel(file);
    }

    /**
     * 使用jackson读取自定义配置文件
     */
    @Test
    void testYaml() {
        File file = new File("C:\\Users\\hujingyi\\Desktop\\test.xlsx");
        jacksonTest.read(file);
    }

    @Test
    void testOOCL() {
        File file = new File("C:\\Users\\hujingyi\\Desktop\\OOCL.xls");
        jacksonTest.read(file);
    }
}
