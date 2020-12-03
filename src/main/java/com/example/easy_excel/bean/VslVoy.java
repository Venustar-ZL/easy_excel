package com.example.easy_excel.bean;

import lombok.Data;
import lombok.ToString;

import java.util.List;

/**
 * @author chenzx
 */
@Data
@ToString
public class VslVoy {

    /**
     * 状态
     */
    private String status;

    /**
     * 船名
     */
    private String vesselName;

    /**
     * 中文船名
     */
    private String vesselChineseName;

    /**
     * 承运人航次
     */
    private String carrierVoyageNo;

    /**
     * 报关航次
     */
    private String terminalVoyageNo;

    /**
     * 船名代码
     */
    private String vesselCode;

    /**
     * 船经人代码
     */
    private String vesselOperatorCode;

    /**
     * 船舶呼号
     */
    private String vesselCallSign;

    /**
     * 船舶IMO编号
     */
    private String vesselImoNo;

    /**
     * 停靠码头代码
     */
    private String portTerminalOfBerthingCode;

    /**
     * 停靠码头名称
     */
    private String portTerminalOfBerthingName;

    /**
     * 船代代码
     */
    private String shippingAgentCode;

    /**
     * 船代名称
     */
    private String shippingAgentName;

    /**
     * 母船代理代码
     */
    private String vesselOperatorAgentCode;

    /**
     * 母船代理名称
     */
    private String vesselOperatorAgentName;

    /**
     * 预计抵港日
     */
    private String eta;

    /**
     * 预计离港日
     */
    private String etd;

    /**
     * 截单时间
     */
    private String siCutOffTime;

    /**
     * VGM截止时间
     */
    private String vgmCutOffTime;

    /**
     * 开港时间
     */
    private String cyOpeningTime;

    /**
     * 截港时间
     */
    private String cyCutOffTime;

    /**
     * 截关时间
     */
    private String etc;

    /**
     * 订舱开始日期
     */
    private String bookingStartTime;

    /**
     * 订舱截止日期
     */
    private String bookingCutOffTime;

    /**
     * 航线Code
     */
    private String lineCode;

    /**
     * 挂靠港列表
     */
    private List<PortOfCall> portOfCalls;

}
