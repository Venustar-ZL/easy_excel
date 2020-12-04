package com.example.easy_excel.config_file_test.config;

import lombok.Data;

/**
 * @Classname VslVoyAttribute
 * @Description TODO
 * @Date 2020/12/3 13:37
 * @Author by ZhangLei
 */
@Data
public class VslVoyAttribute {

    private String name;

    private Integer length;

    /**
     * 圈定可变属性的范围
     */
    private String begin;

    private String end;

}
