package com.microsoul.erp;

import java.util.HashMap;

/**
 * @author tiger (tiger@microsoul.com) 2020/9/29
 */
public class TestPic {
    public static void main(String[] args) {
        HashMap root = new HashMap<>();
        root.put("imageURL","http://www.microsoul.com/images/minishop.png");
        root.put("templatePath","minishop.png");
        Tools.export("./outs", "out_pic.xls", new Object[]{root}, "./tpls/pic");
    }
}
