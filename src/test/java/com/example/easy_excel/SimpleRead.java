package com.example.easy_excel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.example.easy_excel.bean.DemoData;
import com.example.easy_excel.listener.DemoDataListener;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

/**
 * @ClassName:
 * @Description:
 * @Date : 2020-11-30 20:20
 * @Author: ZhangLei
 * @Version: 1.0
 **/
@SpringBootTest
public class SimpleRead {

    /**
     * 第一种写法
     */
    @Test
    public void test1() {
        String fileName = "demo.xlsx";
        EasyExcel.read(fileName, DemoData.class, new DemoDataListener()).sheet().doRead();
    }

    @Test
    public void test2() {
        String fileName = "";
        ExcelReader excelReader = EasyExcel.read(fileName, DemoData.class, new DemoDataListener()).build();
        ReadSheet readSheet = EasyExcel.readSheet(0).build();
        excelReader.read(readSheet);
        excelReader.finish();
    }
}
