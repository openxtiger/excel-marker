package com.microsoul.erp;

import java.util.HashMap;

/**
 * @author tiger (tiger@microsoul.com) 2020/9/29
 */
public class TestSheet {
    public static void main(String[] args) {
        HashMap sheet1 = new HashMap<>();
        sheet1.put("sheet1", "第一个Sheet表格输出的变量");

        HashMap sheet2 = new HashMap<>();
        sheet2.put("sheet2", "第二个Sheet表格输出的内容");

        Tools.export("./outs", "out_sheet.xls", new Object[]{sheet1, sheet2}, "./tpls/sheet");
    }
}
