package com.microsoul.erp.poi;

import com.microsoul.erp.commons.GlobalHelper;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;
import java.util.Stack;

/**
 * 微妞分布式平台-公用工具包
 *
 * @author 广州加叁信息科技有限公司 (tiger@microsoul.com)
 * @version V1.0.0
 */
public class POIParser {


    private HashMap<String, POIProcesser> pMaps;
    private HashMap<String, String> globals;
    private HashMap<String, POIList> plists;
    private POIList clist;

    private Stack<POIList> lists;
    private int listStartRow = Integer.MAX_VALUE;
    private int listEndRow = 0;
    private CreationHelper creationHelper;

    public int parse(Workbook wb, Object obj, int idx) {

        HSSFSheet sheet = (HSSFSheet) wb.getSheetAt(idx);

        HSSFPatriarch patriarch = sheet.getDrawingPatriarch();
        if (patriarch == null) patriarch = sheet.createDrawingPatriarch();

        creationHelper = wb.getCreationHelper();

        pMaps = new HashMap<String, POIProcesser>();
        plists = new HashMap<String, POIList>();
        lists = new Stack<POIList>();
        globals = new HashMap<String, String>();

        POIList root = new POIList();
        root.setName("root");
        root.setSheet(sheet);
        root.setCapacity(1);
        root.setRowStart(1);
        root.setRowEnd(sheet.getLastRowNum() + 1);
        root.setColStart(1);
        root.setColEnd(0);


        lists.push(root);

        //1.parse lists;
        parseList(wb, sheet);

        parseKeys(root, root.getChildren(), obj, wb, patriarch);

        parseVars(root, root.getChildren(), obj, wb, patriarch);

        return 0;
    }

    private void parseKeys(POIList p, ArrayList<POIList> plist, Object obj, Workbook wb, Drawing patriarch) {
        Row sourceRow;
        Cell sourceCell;

        for (int i = p.getRowStart(); i <= p.getRowEnd(); i++) {
            sourceRow = p.getSheet().getRow(i);
            if (sourceRow == null) {
                continue;
            }

            for (int j = p.getColStart(), e = p.getColEnd() == -1 ?
                    sourceRow.getLastCellNum() : p.getColEnd(); j <= e; j++) {
                sourceCell = sourceRow.getCell(j);
                if (sourceCell == null) {
                    continue;
                }

                makeCellStyle(sourceCell.getCellStyle(), sourceCell, obj, wb);

                if (sourceCell.getCellType() != CellType.STRING) continue;
                String v = sourceCell.getRichStringCellValue().getString().trim();
                if (v.contains("${%")) {
                    //System.out.println(parseV(v));
                    sourceCell.setCellValue(parseV(v));
                }


            }
        }


    }

    private String parseV(String v) {
        int idx = v.indexOf("${%"), idx2;
        String s;
        if (idx >= 0) {
            idx2 = v.indexOf("}", idx);
            if (idx2 == -1) return v;

            String[] vs = v.substring(idx + 3, idx2).split(",");

            s = globals.get(vs[0]);
            if (s == null) {
                return v;
            }

            for (int k = 1; k < vs.length; k += 2) {
                s = s.replaceAll(vs[k], vs[k + 1]);
            }

            return parseV(v.substring(0, idx) + s + v.substring(idx2 + 1));

        }
        return v;
    }

    private void parseVars(POIList p, ArrayList<POIList> plist, Object obj, Workbook wb, Drawing patriarch) {
        for (POIList poi : plist) {
            parseVars(poi, poi.getChildren(), obj, wb, patriarch);
        }
        Row sourceRow;
        Cell sourceCell;
        POIProcesser processer = new POIProcesser();

        processer.init(obj, p.getSheet(), wb, patriarch, p);

        p.setProcesser(processer);


        if (p.getParent() == null || p.getParent().getName().equals("root")) {
            pMaps.put(processer.getName(), processer);
        } else {
            pMaps.put(p.getParent().getName() + "." + processer.getName(), processer);
        }


        for (int i = p.getRowStart(); i <= p.getRowEnd(); i++) {
            sourceRow = p.getSheet().getRow(i);
            if (sourceRow == null) {
                continue;
            }

            for (int j = p.getColStart(), e = p.getColEnd() == -1 ?
                    sourceRow.getLastCellNum() : p.getColEnd(); j <= e; j++) {
                sourceCell = sourceRow.getCell(j);
                if (sourceCell == null) {
                    continue;
                }

                makeCellStyle(sourceCell.getCellStyle(), sourceCell, obj, wb);

                if (sourceCell.getCellType() != CellType.STRING) continue;
                String v = sourceCell.getRichStringCellValue().getString().trim();
                makeKey(v, processer, p, sourceRow, sourceCell);
            }
        }

        if (p.getColumns() > 0) {

            for (int l = p.getColumns(), i = 0; i < l; i++) {
                POIProcesser sprocesser = new POIProcesser();
                sprocesser.init(obj, p.getSheet(), wb, patriarch, p);
                sprocesser.setColOffset(i * p.getColOffset());
                processer.addSibling(sprocesser);

            }
        }
    }

    private void makeCellStyle(CellStyle style, Cell targetCell, Object obj, Workbook wb) {
        String v = style.getDataFormatString();
        v = v.replaceAll("\\\\", "");

        ArrayList<String> vs = GlobalHelper.split("~([^\\}]*)~", v);
        if (vs.size() == 1) return;
        StringBuilder b = new StringBuilder();
        if (obj instanceof Map) {
            Map o = (Map) obj;
            int i, l = vs.size();
            for (i = 0; i < l / 2; i += 2) {
                b.append(vs.get(i));
                b.append(o.get(vs.get(i + 1)));
            }
            if (i < l) {
                b.append(vs.get(l - 1));
            }
            style.setDataFormat(wb.createDataFormat().getFormat(b.toString()));
            targetCell.setCellStyle(style);
        }

    }

    /**
     * Analysis key
     * ${key}
     * $=     _FORMULA
     * $~()   _PICTURE
     * $~picture(col_d,row_d,px1,py1,px2,py2) set the picture
     * $~!picture(px1,py1,width,height) set the picture with the width and height
     * $~#picture(height,showText==1,bartype=ean13|code128,px1,py1,type,width,height) set the bar code picture with the width and height
     * $~&picture_path(px1,py1,type,width,height) set the picture by picture_path for tpl path
     * type=0,with width and height
     * type=1,with width
     * type=2,with height
     * type=3,if height is null then within width else within width and height
     * <p/>
     * $!     _EVAL
     * <p/>
     * ${$name.var}   find by name's object's var
     * ${$var}   find by root's object's var
     * <p/>
     * ${#name.var}   find by name's processer's var
     * ${#var}   find by self's processer's var
     * <p/>
     * ${!var}   find by var's var.
     * ${!}   get the curr row number
     * ${?}   get the curr col number
     * <p/>
     * ${^indexvar~array1~array2~...}   find by array
     * <p/>
     * ${#name.:COL+-0}   COL{start}:COL{end}
     * <p/>
     * ${#name.&COL+-0}   COL{begin1},COL{begin2},...
     * <p/>
     * ${#name.$COL+-0}   COL{end1},COL{end2},...
     * <p/>
     * ${#name.#FUN+-0}   FUN(CurrCOL{start}:CurrCOL{end})
     * <p/>
     * ${#name.#+-0}   SUM(CurrCOL{start}:CurrCOL{end})
     * <p/>
     * $:name+-0   SUM(CurrCOL{start}:CurrCOL{end})
     * <p/>
     * $:name+-0:COL   SUM(COL{start}:COL{end}) => $=${#name.#COL+-0} => $=SUM(${#name.:COL+-0})
     * <p/>
     * $$name+-0   SUM(CurrCOL{end1},CurrCOL{end2})
     * $&name+-0   SUM(CurrCOL{being1},CurrCOL{begin2})
     * <p/>
     * $$name+COL+-0  => $=SUM(${#name.$COL+-0})
     * $&name+COL+-0  => $=SUM(${#name.&COL+-0})
     * <p/>
     * <p/>
     * #{key} numeric key
     * ${key($name,prex,startIndex)} get pollist's vars,
     *
     * @param v
     * @param p
     * @param sourceRow
     * @param targetCell
     * @return
     */
    private POIProcesser makeKey(String v, POIProcesser p, POIList plist, Row sourceRow, Cell targetCell) {

        ArrayList<String> vs = GlobalHelper.split("[\\$\\#]\\{([^\\}]*)\\}", v);

        if (vs.size() > 0) {
            String f = vs.get(0);
            int style = 0;
            int pstyle = POIVar.TYPE_PICTURE1;
            boolean iskey = false;
            if (f.startsWith("$") && f.length() >= 2) {
                iskey = true;
                switch (f.charAt(1)) {
                    case '=':
                        f = f.substring(2);
                        style = POIVar.TYPE_FORMULA;
                        break;
                    case ':':
                        style = POIVar.TYPE_FORMULA; //
                        vs.add(f.substring(1));
                        f = "";
                        break;
                    case '$':
                    case '&':
                        String b;
                        style = POIVar.TYPE_FORMULA;
                        b = f.substring(2);
                        int idx = b.indexOf("+");
                        if (idx > 0) {
                            b = b.substring(0, idx).trim() + "." + f.charAt(1) + "+" + b.substring(idx + 1);
                        } else {
                            b = b.trim() + "." + f.charAt(1);
                        }

                        vs.add("#" + b); //#list +-1  //#list.$+-1
                        vs.add(")");
                        f = "SUM(";
                        break;
                    case '~':
                        pstyle = POIVar.TYPE_PICTURE;
                    case '`':
                        f = f.substring(2);
                        String[] fs = f.split("[\\(\\),]");
                        if (fs.length == 0) return p;

                        POIVar c = new POIVar();
                        c.setStyle(pstyle);
                        c.setInitTargetCell(targetCell);
                        c.setStrings(fs);
                        c.setVariables(new ArrayString[]{new ArrayString(fs)});
                        p.addVar(c);
                        //checkFixedVar(plist, c);

                        targetCell.setCellValue(creationHelper.createRichTextString(""));
                        return p;
                    case '!':
                        f = f.substring(2);
                        style = POIVar.TYPE_EVAL;
                        break;
                }
            }

            if (vs.size() == 1 && !iskey)
                return p;

            if (v.trim().startsWith("#")) {
                style = POIVar.TYPE_NUMERIC;
            }
            POIVar c = new POIVar();

            c.setStyle(style);
            c.setInitTargetCell(targetCell);
            //checkFixedVar(plist, c);

            ArrayList<String> strs = new ArrayList<String>();
            ArrayList<ArrayString> vars = new ArrayList<ArrayString>();

            strs.add(f);

            for (int i = 2, l = vs.size(); i < l; i += 2) {
                f = vs.get(i);
                strs.add(f);
            }
            String[] ss = new String[strs.size()];
            c.setStrings(strs.toArray(ss));

            for (int i = 1, l = vs.size(); i < l; i += 2) {
                ArrayString s = new ArrayString(vs.get(i).split("[\\(\\),\\]]"));
                vars.add(s);
            }
            ArrayString[] as = new ArrayString[vars.size()];
            c.setVariables(vars.toArray(as));

            p.addVar(c);
            targetCell.setCellValue(creationHelper.createRichTextString(""));

        }
        return p;
    }

    private void checkFixedVar(POIList plist, POIVar c) {
        int row = c.getTargetCell().getRowIndex();
        for (POIList p : plist.getChildren()) {
            if (row > p.getRowEnd()) {
                c.setIsFixed(true);
                break;
            }
        }
    }

    private void parseList(Workbook wb, Sheet sheet) {
        int pStartRow = sheet.getFirstRowNum(), pEndRow = sheet.getLastRowNum();
        Row sourceRow;
        Cell sourceCell;
        int r, idx;
        for (int i = pStartRow; i <= pEndRow; i++) {
            sourceRow = sheet.getRow(i);
            if (sourceRow == null) {
                continue;
            }
            for (int j = sourceRow.getFirstCellNum(), k = sourceRow.getLastCellNum(); j <= k; j++) {

                sourceCell = sourceRow.getCell(j);
                if (sourceCell == null) {
                    continue;
                }
                if (sourceCell.getCellType() != CellType.STRING) continue;
                String v = sourceCell.getRichStringCellValue().getString().trim();
                makeListKey(v, i, sheet, sourceCell);
                makeGlobalKey(v, i, sheet, sourceCell);
            }

        }

        //pasrse list vars  index start with row_number
        //${{name.prex
        // ... offset
        // key=>value
        // value
        //$}}
        for (int i = pStartRow; i <= pEndRow; i++) {
            sourceRow = sheet.getRow(i);
            if (sourceRow == null) {
                continue;
            }
            for (int j = sourceRow.getFirstCellNum(), k = sourceRow.getLastCellNum(); j <= k; j++) {

                sourceCell = sourceRow.getCell(j);

                if (sourceCell == null) {
                    continue;
                }
                if (sourceCell.getCellType() != CellType.STRING) continue;
                String v = sourceCell.getRichStringCellValue().getString().trim();

                if (v.startsWith("${{")) {
                    sourceCell.setCellValue(creationHelper.createRichTextString(""));
                    idx = v.lastIndexOf(".");
                    String prex = "";
                    if (idx > 0) {
                        prex = v.substring(idx + 1);
                        v = v.substring(0, idx);
                    }
                    clist = plists.get(v.substring(3));

                    if (clist != null) {
                        int offset = 0;
                        for (++i; i <= pEndRow; i++) {
                            sourceRow = sheet.getRow(i);
                            if (sourceRow == null) {
                                clist.put(prex + String.valueOf(i + 1 + offset), "");
                                continue;
                            }
                            sourceCell = sourceRow.getCell(j);
                            if (sourceCell == null) {
                                clist.put(prex + String.valueOf(i + 1 + offset), "");
                                continue;
                            }
                            v = sourceCell.getRichStringCellValue().getString().trim();
                            if (v.startsWith("$}}")) {
                                sourceCell.setCellValue(creationHelper.createRichTextString(""));
                                clist = null;
                                break;
                            }
                            if (v.startsWith("...")) {
                                sourceCell.setCellValue(creationHelper.createRichTextString(""));
                                int c = GlobalHelper.parseInt(v.substring(3), 0);
                                offset += c - 1;
                                continue;
                            }
                            if ((idx = v.indexOf("=>")) > 0) {
                                clist.put(v.substring(0, idx), v.substring(idx + 2));
                            } else {
                                clist.put(prex + String.valueOf(i + 1 + offset), v);
                            }
                            sourceCell.setCellValue(creationHelper.createRichTextString(""));
                        }
                    }
                }
            }
        }
    }


    /**
     * external list
     * start tag: ${{list
     * end tag:   $}}
     *
     * @param v     row value
     * @param row   row
     * @param sheet sheet
     * @param cell  cell
     */

    private void makeListKey(String v, int row, Sheet sheet, Cell cell) {
        ArrayList<String> vs = GlobalHelper.split("\\$([<>]{2}.*)", v);   //
        POIList c;

        if (vs.size() > 1) {
            String f = vs.get(1);
            switch (f.charAt(1)) {
                case '<':
                    f = f.substring(2);
                    String[] fun = f.split("[\\(\\),]");
                    if (fun.length < 5) break;
                    c = new POIList();
                    c.setSheet(sheet);
                    c.setParent(lists.peek());
                    c.setName(fun[0]);
                    c.setRowStart(GlobalHelper.parseInt(fun[1], 0));
                    c.setRowEnd(GlobalHelper.parseInt(fun[2], 0));
                    c.setColStart(GlobalHelper.parseInt(fun[3], 0));
                    c.setColEnd(GlobalHelper.parseInt(fun[4], 0));
                    /*c.setDataRows(com.microsoul.weapi.util.WeosoGlobals.parseInt(fun[5], 0));*/
                    c.setCapacity(GlobalHelper.parseInt(fun[5], 0));
                    c.setPageOffset(fun.length >= 7 ? GlobalHelper.parseInt(fun[6], 0) : 0);
                    c.setPageRows(fun.length >= 8 ? GlobalHelper.parseInt(fun[7], 0) : 0);
                    c.setColOffset(fun.length >= 9 ? GlobalHelper.parseInt(fun[8], 0) : 0);
                    c.setColumns(fun.length >= 10 ? GlobalHelper.parseInt(fun[9], 0) : 0);

                    lists.push(c);
                    plists.put(fun[0], c);
                    break;
                case '>':
                    lists.pop();
                    break;
            }
            cell.setCellValue(creationHelper.createRichTextString(""));

            listStartRow = Math.min(listStartRow, row);
            listEndRow = Math.max(listEndRow, row);
        }
    }

    private void makeGlobalKey(String v, int row, Sheet sheet, Cell cell) {
        int i;

        if (v != null && v.startsWith("$[[") && v.endsWith("]]")) {
            v = v.substring(3, v.length() - 2);
            i = v.indexOf("=>");
            globals.put(v.substring(0, i), v.substring(i + 2));
            cell.setCellValue(creationHelper.createRichTextString(""));
        }

    }

    public int process(Object obj, File pf, int insidePictureCount) {
        POIProcesser processer = pMaps.get("root");
        processer.process(obj, pMaps, pf, insidePictureCount);


        return 0;
    }


}
