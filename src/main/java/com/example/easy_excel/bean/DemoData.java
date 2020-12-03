package com.example.easy_excel.bean;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;


/**
 * @ClassName:
 * @Description:
 * @Date : 2020-11-30 20:02
 * @Author: ZhangLei
 * @Version: 1.0
 **/
@Data
public class DemoData {

    @ExcelProperty(value = "学号")
    private Integer id;

    @ExcelProperty(value = "姓名")
    private String name;

    @ExcelProperty(value = "性别")
    private String sex;

    @ExcelProperty(value = "年龄")
    private Integer age;

}
