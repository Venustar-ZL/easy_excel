package com.example.easy_excel.config_file_test.config;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.annotation.Configuration;

/**
 * @Classname ConfigBean
 * @Description TODO
 * @Date 2020/12/2 13:37
 * @Author by ZhangLei
 */
@Configuration
@Data
@AllArgsConstructor
@NoArgsConstructor
public class CNCConfigBean {

    @Value("${cnc.head.accurate}")
    private Boolean headAccurate;

    @Value("${cnc.head.color}")
    private String headColor;

    @Value("${cnc.head.pattern}")
    private String headPattern;

    @Value("${cnc.head.parsing}")
    private Boolean headParsing;

    @Value("${cnc.head.range}")
    private Integer headRange;

    @Value("${cnc.tail.accurate}")
    private Boolean tailAccurate;

    @Value("${cnc.tail.color}")
    private String tailColor;

    @Value("${cnc.tail.pattern}")
    private String tailPattern;

    @Value("${cnc.tail.parsing}")
    private Boolean tailParsing;

    @Value("${cnc.tail.range}")
    private Integer tailRange;

}
