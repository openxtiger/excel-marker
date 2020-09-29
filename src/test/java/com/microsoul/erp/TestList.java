package com.microsoul.erp;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author tiger (tiger@microsoul.com) 2020/9/29
 */
public class TestList {
    public static void main(String[] args) {
        HashMap root = new HashMap<>();
        root.put("title", "字符串");
        root.put("price", 12.345);

        List list = new ArrayList();

        Map listObj = new HashMap();
        listObj.put("name", "张三");
        listObj.put("age", 31);
        list.add(listObj);

        listObj = new HashMap();
        listObj.put("name", "李四");
        listObj.put("age", 30);
        list.add(listObj);

        listObj = new HashMap();
        listObj.put("name", "王五");
        listObj.put("age", 29);
        list.add(listObj);


        root.put("list", list);

        Tools.export("./outs", "out_list.xls", new Object[]{root}, "./tpls/list");
    }
}
