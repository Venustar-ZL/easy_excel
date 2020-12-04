package com.example.easy_excel.bean;

import lombok.Data;

/**
 * @author chenzx
 */
@Data
public class PortOfCall {

    /**
     * 挂靠港序号
     */
    private String portOfCallNo;

    /**
     * 挂靠港代码
     */
    private String portOfCallCode;

    /**
     * 挂靠港名称
     */
    private String portOfCallName;

    /**
     * 预计抵港日
     */
    private String eta;
}
