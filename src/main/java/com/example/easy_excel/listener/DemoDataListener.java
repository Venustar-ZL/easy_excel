package com.example.easy_excel.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.fastjson.JSONObject;
import com.example.easy_excel.bean.DemoData;
import lombok.extern.slf4j.Slf4j;

import java.util.ArrayList;
import java.util.List;

/**
 * @ClassName:
 * @Description:
 * @Date : 2020-11-30 20:06
 * @Author: ZhangLei
 * @Version: 1.0
 **/
@Slf4j
public class DemoDataListener extends AnalysisEventListener<DemoData> {

    List<DemoData> list = new ArrayList<>();

    /**
     * 如果使用了spring，请使用这个构造方法。每次创建Listener的时候需要把spring管理的类传进来
     */
    public DemoDataListener() {}

    /**
     * 这里的每一条数据解析都会来调用
     * @param demoData
     * @param analysisContext
     */
    @Override
    public void invoke(DemoData demoData, AnalysisContext analysisContext) {

        log.info("解析到一条数据：{}", JSONObject.toJSONString(demoData));
        list.add(demoData);

    }

    /**
     * 所有的数据解析完成后都会调用
     * @param analysisContext
     */
    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {

        log.info("解析完成后数据：{}", JSONObject.toJSONString(list));

    }
}
