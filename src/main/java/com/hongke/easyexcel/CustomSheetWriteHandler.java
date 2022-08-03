package com.hongke.easyexcel;

import com.alibaba.excel.write.handler.SheetWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Arrays;
import java.util.Random;

/**
 * 自定义拦截器.对第一列第一行和第二行的数据新增下拉框，显示 测试1 测试2
 *
 * @author Jiaju Zhuang
 */
public class CustomSheetWriteHandler implements SheetWriteHandler {

    private static final Logger LOGGER = LoggerFactory.getLogger(CustomSheetWriteHandler.class);

    @Override
    public void beforeSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {

    }

    @Override
    public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
        LOGGER.info("第{}个Sheet写入成功。", writeSheetHolder.getSheetNo());

        // 区间设置 第一列第一行和第二行的数据。由于第一行是头，所以第一、二行的数据实际上是第二三行
        CellRangeAddressList cellRangeAddressList = new CellRangeAddressList(1, 3000, 0, 0);
        DataValidationHelper helper = writeSheetHolder.getSheet().getDataValidationHelper();
        DataValidationConstraint constraint = helper.createExplicitListConstraint(new String[]{"111"});
        DataValidation dataValidation = helper.createValidation(constraint, cellRangeAddressList);
        writeSheetHolder.getSheet().addValidationData(dataValidation);

        Workbook workbook = writeWorkbookHolder.getWorkbook();
        String sheetName = "SheetHidden";
        // 1.创建一个隐藏的sheet 名称为 sheet
        Sheet sheet = workbook.createSheet(sheetName);
        // 设置隐藏
        workbook.setSheetHidden(workbook.getSheetIndex(sheetName), true);
        String[] arr = getBatchArr("dcp");
        for (int i = 0; i < arr.length; i++) {
            // i:表示你开始的行数 0表示你开始的列数
            sheet.createRow(i).createCell(0).setCellValue(arr[i]);
        }
        Name category1Name = workbook.createName();
        category1Name.setNameName(sheetName);
        // 4 $A$1:$A$N代表 以A列1行开始获取N行下拉数据
        category1Name.setRefersToFormula(sheetName + "!$A$1:$A$" + (arr.length));
        // 5 将刚才设置的sheet引用到你的下拉列表中 //起始行、终止行、起始列、终止列
        CellRangeAddressList addressList = new CellRangeAddressList(1, 65535, 2, 2);
        DataValidationConstraint constraint8 = helper.createFormulaListConstraint(sheetName);
        DataValidation dataValidation3 = helper.createValidation(constraint8, addressList);
        writeSheetHolder.getSheet().addValidationData(dataValidation3);

        String sheetName1 = "SheetHidden1";
        sheet(sheetName1, workbook, getBatchArr("mdp"));
        // 5 将刚才设置的sheet引用到你的下拉列表中 //起始行、终止行、起始列、终止列
        CellRangeAddressList addressList1 = new CellRangeAddressList(1, 65535, 3, 3);
        DataValidationConstraint constraint81 = helper.createFormulaListConstraint(sheetName1);
        DataValidation dataValidation31 = helper.createValidation(constraint81, addressList1);
        writeSheetHolder.getSheet().addValidationData(dataValidation31);
    }

    private void sheet(String sheetName1, Workbook workbook, String[] arr) {
        // 1.创建一个隐藏的sheet 名称为 sheet
        Sheet sheet1 = workbook.createSheet(sheetName1);
        // 设置隐藏
        workbook.setSheetHidden(workbook.getSheetIndex(sheetName1), true);
        for (int i = 0; i < arr.length; i++) {
            // i:表示你开始的行数 0表示你开始的列数
            sheet1.createRow(i).createCell(0).setCellValue(arr[i]);
        }
        Name category1Name1 = workbook.createName();
        category1Name1.setNameName(sheetName1);
        // 4 $A$1:$A$N代表 以A列1行开始获取N行下拉数据
        category1Name1.setRefersToFormula(sheetName1 + "!$A$1:$A$" + (arr.length));
    }

    private String[] getBatchArr(String suffix) {
        String[] arr = new String[1000];
        for(int i = 0; i < arr.length; i++) {
            arr[i] = new Random().nextInt() + suffix;
        }
        return arr;
    }
}
