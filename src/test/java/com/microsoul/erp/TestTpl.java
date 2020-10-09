package com.microsoul.erp;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author tiger (tiger@microsoul.com) 2020/9/29
 */
public class TestTpl {
    public static void main(String[] args) {
        HashMap root = new HashMap<>();
        root.put("code","SL0000001");



        Tools.export("./outs", "out_tpl.xls", new Object[]{root}, "./tpls/tpl");
    }
}
