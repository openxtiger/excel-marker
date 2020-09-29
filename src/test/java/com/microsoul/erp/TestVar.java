package com.microsoul.erp;

import java.util.HashMap;

/**
 * @author tiger (tiger@microsoul.com) 2020/9/29
 */
public class TestVar {
    public static void main(String[] args) {
        HashMap root = new HashMap<>();
        root.put("title", "字符串");
        root.put("price", 12.345);
        Tools.export("./outs", "out_var.xls", new Object[]{root}, "./tpls/var");
    }
}
