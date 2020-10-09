package com.microsoul.erp.poi;

import org.apache.poi.ss.usermodel.Cell;

/**
 * 微妞分布式平台-公用工具包
 * @author 广州加叁信息科技有限公司 (tiger@microsoul.com)
 * @version V1.0.0
 */
public class POIVar {
    public static final int TYPE_PICTURE = 1;
    public static final int TYPE_PICTURE1 = -1;
    public static final int TYPE_FORMULA = 2;
    public static final int TYPE_EVAL = 3;
    public static final int TYPE_NUMERIC = 4;
    public static final int TYPE_STYLE = 5;
    public static final int TYPE_ARRAY = 6;
    private String[] strings;
    private ArrayString[] variables;
    private Cell targetCell;
    private Cell initTargetCell;
    private Cell lastTargetCell;
    private int style;
    private Object obj;
    private boolean isFixed = false;

    public void setStrings(String[] strs) {
        this.strings = strs;
    }

    public void setVariables(ArrayString[] vars) {
        this.variables = vars;
    }

    public void setInitTargetCell(Cell targetCell) {
        this.initTargetCell = targetCell;
        this.targetCell = targetCell;
        this.lastTargetCell = targetCell;
    }

    public void setTargetCell(Cell targetCell) {
        this.targetCell = targetCell;
        this.lastTargetCell = targetCell;
    }

    public POIVar reset() {
        this.targetCell = initTargetCell;
        this.lastTargetCell = initTargetCell;
        return this;
    }

    public void setLastTargetCell(Cell lastTargetCell) {
        this.lastTargetCell = lastTargetCell;
    }

    public String[] getStrings() {
        return strings;
    }

    public ArrayString[] getVariables() {
        return variables;
    }

    public void setStyle(int style) {
        this.style = style;
    }

    public int getStyle() {
        return style;
    }

    public Cell getTargetCell() {
        return targetCell;
    }

    public Cell getLastTargetCell() {
        return lastTargetCell;
    }

    public Object getObj() {
        return obj;
    }

    public void setObj(Object obj) {
        this.obj = obj;
    }

    public void setIsFixed(boolean fixed) {
        isFixed = fixed;
    }

    public boolean isFixed() {
        return isFixed;
    }
}

