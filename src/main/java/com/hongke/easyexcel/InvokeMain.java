package com.hongke.easyexcel;

import com.alibaba.excel.EasyExcel;

import java.util.ArrayList;
import java.util.List;

public class InvokeMain {

    public static void main(String[] args) {
        List<DemoData> demoDataList = new ArrayList<>();
        DemoData demoData = new DemoData();
        demoData.setString("哈哈哈");
        demoDataList.add(demoData);
        String fileName = "D:\\2_work_for_own\\projects-open\\easyexcel-dropdown\\file\\1.xlsx";
        EasyExcel.write(fileName, DemoData.class)
                .registerWriteHandler(new CustomSheetWriteHandler())
                .sheet("模板").doWrite( demoDataList);
    }

}
