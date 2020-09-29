package com.microsoul.erp.poi;

import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;
import java.util.HashMap;

/**
 * 微妞分布式平台-公用工具包
 * @author 广州加叁信息科技有限公司 (tiger@microsoul.com)
 * @version V1.0.0
 */
public class POIList {

    private Sheet sheet;
    private String name;
    private POIList parent;
    private int rowStart;
    private int initRowStart;
    private int rowEnd;
    private int initRowEnd;
    private int colStart;
    private int colEnd;
    private int initColEnd;
    private int initColStart;

    private ArrayList<POIList> children = new ArrayList<POIList>();
    private HashMap<String, String> vars = new HashMap<String, String>();

    //private int dataRows;
    private POIProcesser processer;
    private int capacity;
    private int realDataRows;
    private int offsets = 0;
    private int pageOffset;
    private int pageRows;
    private boolean count2Page;
    private int colOffset;
    private int columns;


    public void setSheet(Sheet sheet) {
        this.sheet = sheet;
    }

    public void setParent(POIList parent) {
        this.parent = parent;
        this.parent.add(this);
    }

    public POIList getParent() {
        return parent;
    }

    private void add(POIList c) {
        this.children.add(c);
    }

    public int getRealDataRows() {
        return realDataRows == 0 ? getDataRows() : realDataRows;
    }

    public void resetRealDataRows(int offset) {
        this.realDataRows = getDataRows() + offset;
        if (offset > 0)
            this.offsets += offset;
    }

    public int getOffsets() {
        return offsets;
    }

    public int getRowStart() {
        return rowStart;
    }


    public void setRowStart(int rowStart) {
        this.rowStart = rowStart - 1;
        this.initRowStart = rowStart - 1;
    }

    public int getRowEnd() {
        return rowEnd;
    }

    public void setRowEnd(int rowEnd) {
        this.rowEnd = rowEnd - 1;
        this.initRowEnd = rowEnd - 1;
    }

    public int getInitRowStart() {
        return initRowStart;
    }

    public int getInitRowEnd() {
        return initRowEnd;
    }

    public int getColStart() {
        return colStart;
    }

    public void setColStart(int colStart) {
        this.colStart = colStart - 1;
        this.initColStart = colStart - 1;
    }

    public int getColEnd() {
        return colEnd;
    }

    public void setColEnd(int colEnd) {
        this.colEnd = colEnd - 1;
        this.initColEnd = colEnd - 1;
    }

    public int getDataRows() {
        return rowEnd - rowStart + 1;
    }

    public int getPageOffset() {
        return pageOffset;
    }

    public void setPageOffset(int pageOffset) {
        this.pageOffset = pageOffset;
    }

    public int getPageCount(int count) {
        return count == 0 ? 1 : count * (this.rowEnd - this.rowStart + 1) / (pageRows == 0 ? 1 : pageRows);
    }

    public int getPageRows() {
        return pageRows;
    }

    public void setPageRows(int pageRows) {
        if (pageRows < 0) {
            count2Page = true;
            this.pageRows = -pageRows;
            return;
        }
        this.pageRows = pageRows * (this.rowEnd - this.rowStart + 1);
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public ArrayList<POIList> getChildren() {
        return children;
    }

    public Sheet getSheet() {
        return sheet;
    }

    public POIProcesser getProcesser() {
        return processer;
    }

    public void setProcesser(POIProcesser processer) {
        this.processer = processer;
    }

    public void setCapacity(int capacity) {
        this.capacity = capacity;
    }

    public int getCapacity() {
        return capacity;
    }

    public void resetOffset(int colOffset) {
        this.rowStart = initRowStart;
        this.rowEnd = initRowEnd;
        this.colStart = initColStart + colOffset;
        this.colEnd = initColEnd + colOffset;
        this.pageOffset = -1;
    }

    public void reset(int rows) {
        this.rowStart += rows;
        this.rowEnd += rows;

    }

    public void shiftRows(int rows) {
        this.rowStart += rows;
        this.rowEnd += rows;
        for (POIList l : children) {
            l.shiftRows(rows);
        }
    }

    @Override
    public String toString() {
        return "POIList{" +
                "name='" + name + '\'' +
                ", parent=" + parent +
                ", rowStart=" + rowStart +
                ", rowEnd=" + rowEnd +
                ", colStart=" + colStart +
                ", colEnd=" + colEnd +
                '}';
    }


    public void put(String key, String value) {
        vars.put(key, value);
    }

    public String get(String key) {
        return vars.get(key);
    }

    public boolean isCount2Page() {
        return count2Page;
    }

    public void setColOffset(int colOffset) {
        this.colOffset = colOffset;
    }

    public int getColOffset() {
        return colOffset;
    }

    public void setColumns(int columns) {
        this.columns = columns;
    }

    public int getColumns() {
        return columns;
    }
}
