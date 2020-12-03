package com.example.easy_excel.config_file_test.config;


import lombok.Data;
import lombok.ToString;

/**
 * @Classname ConfigBean
 * @Description TODO
 * @Date 2020/12/2 18:19
 * @Author by ZhangLei
 */
@Data
@ToString
public class ConfigBean {

    private Head head;

    private Content content;

    private Tail tail;

}
