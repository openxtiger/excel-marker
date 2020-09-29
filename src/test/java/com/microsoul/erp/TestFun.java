package com.microsoul.erp;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author tiger (tiger@microsoul.com) 2020/9/29
 */
public class TestFun {
    public static void main(String[] args) {
        HashMap root = new HashMap<>();
        root.put("discount",0.9);

        List list = new ArrayList();

        Map listObj = new HashMap();
        listObj.put("name", "苹果");
        listObj.put("price", 12.5);
        listObj.put("qty", 20);
        list.add(listObj);

        listObj = new HashMap();
        listObj.put("name", "梨");
        listObj.put("price", 19.5);
        listObj.put("qty", 20);
        list.add(listObj);

        listObj = new HashMap();
        listObj.put("name", "桃子");
        listObj.put("price", 18.5);
        listObj.put("qty", 20);
        list.add(listObj);


        root.put("list", list);

        Tools.export("./outs", "out_fun.xls", new Object[]{root}, "./tpls/fun");
    }
}
