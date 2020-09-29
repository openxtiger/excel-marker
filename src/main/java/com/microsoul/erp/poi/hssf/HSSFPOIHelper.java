package com.microsoul.erp.poi.hssf;

import com.microsoul.erp.poi.POIHelper;
import org.apache.poi.hssf.usermodel.HSSFEvaluationWorkbook;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.FormulaParser;
import org.apache.poi.ss.formula.FormulaRenderer;
import org.apache.poi.ss.formula.FormulaType;
import org.apache.poi.ss.formula.ptg.Ptg;
import org.apache.poi.ss.formula.ptg.RefPtg;
import org.apache.poi.ss.usermodel.*;

/**
 * 微妞分布式平台-公用工具包
 *
 * @author 广州加叁信息科技有限公司 (tiger@microsoul.com)
 * @version V1.0.0
 */
public class HSSFPOIHelper implements POIHelper {
    private Workbook workbook;
    private CreationHelper creationHelper;
    private Sheet sheet;
    private Drawing drawing;

    public HSSFPOIHelper(Workbook workbook, Sheet sheet, Drawing drawing) {
        this.workbook = workbook;
        this.sheet = sheet;
        this.drawing = drawing;
        this.creationHelper = workbook.getCreationHelper();

    }


    public void copyCell(int rowStart, int colStart, int colEnd, int dataRows, int offset, int count) {
        Row r;

        Row targetRow;
        Cell sourceCell, targetCell;
        for (int c = 1; c <= count; c++) {
            for (int i = rowStart; i < rowStart + dataRows; i++) {
                r = sheet.getRow(i);
                if (r == null) continue;
                targetRow = sheet.getRow(i + offset + dataRows * (c - 1));

                if (targetRow == null)
                    targetRow = sheet.createRow(i + offset + dataRows * (c - 1));


                for (int j = colStart; j <= colEnd; j++) {
                    sourceCell = r.getCell(j);
                    if (sourceCell == null) {
                        continue;
                    }
                    targetCell = targetRow.getCell(j);
                    if (targetCell == null)
                        targetCell = targetRow.createCell(j);

                    targetCell.setCellStyle(sourceCell.getCellStyle());

                    switch (sourceCell.getCellType()) {
                        case BOOLEAN:
                            targetCell.setCellValue(sourceCell.getBooleanCellValue());
                            break;
                        case ERROR:
                            targetCell.setCellErrorValue(sourceCell.getErrorCellValue());
                            break;
                        case FORMULA:
                            parseFormula(sourceCell, targetCell, rowStart, rowStart + dataRows);
                            break;
                        case NUMERIC:
                            targetCell.setCellValue(sourceCell.getNumericCellValue());
                            break;
                        case STRING:
                            targetCell.setCellValue(sourceCell.getRichStringCellValue());
                            break;
                    }
                }
            }
        }
    }

    private void parseFormula(Cell sourceCell, Cell targetCell, int startRow, int endRow) {

        Ptg[] ptgs = FormulaParser.parse(sourceCell.getCellFormula(),
                HSSFEvaluationWorkbook.create((HSSFWorkbook) workbook),
                FormulaType.CELL, -1);

        boolean changed = false;
        //int row = sourceCell.getRowIndex();

        for (Ptg ptg : ptgs) {
            if (ptg instanceof RefPtg) {
                RefPtg rptg = (RefPtg) ptg;
                //System.out.println("----------------------" + sourceCell + "," + rptg.toFormulaString() + "," + targetCell);
                if (rptg.getRow() >= startRow && rptg.getRow() <= endRow) {
                    rptg.setRow(rptg.getRow() + targetCell.getRowIndex() - sourceCell.getRowIndex());
                    changed = true;
                }
            }
        }
        if (changed) {
            targetCell.setCellFormula(FormulaRenderer.toFormulaString(HSSFEvaluationWorkbook.create((HSSFWorkbook) workbook), ptgs));
        } else {
            targetCell.setCellFormula(sourceCell.getCellFormula());
        }
    }

    public void createPicture(ClientAnchor anchor,
                              int width, int height,
                              byte[] pictureData,
                              int insidePictureCount) {

        int col2 = anchor.getCol2();
        int col = anchor.getCol1();
        int row2 = anchor.getRow2();
        int row = anchor.getRow1();
        int px1 = anchor.getDx1();
        int px2 = anchor.getDx2();
        int py1 = anchor.getDy1();
        int py2 = anchor.getDy2();

        double w = 0;
        for (int i = col2; i >= col; i--) {
            w += getColumnWidthInPixels(sheet, i);
        }


        double h = 0;
        for (int i = row2; i >= row; i--) {
            h += getRowHeightInPixels(sheet, i);
        }

        if (px1 + px2 == 0) {
            w -= 10;
        }
        if (py1 + py2 == 0) {
            h -= 10;
        }
        double scale = Math.min(w / width, h / height);
        double scaledWidth = width, scaledHeight = height;
        if (scale < 1) {
            scaledWidth = width * scale;
            scaledHeight = height * scale;
        }
        //System.out.println("w=" + w + "," + "width=" + width + "," + "scaledWidth=" + scaledWidth);
        if (px1 + px2 == 0) {  //center
            if (w > scaledWidth) {
                px1 = (int) (w + 10 - scaledWidth) / 2;
            } else if (w > width) {
                px1 = (int) (w + 10 - width) / 2;
            } else {
                px1 = 5;
            }
        } else {
            w -= px1 + px2;
        }
        //System.out.println("h=" + h + "," + "height=" + height + "," + "scaledHeight=" + scaledHeight + "\n" + px1 + "," + py1 + "|" + scale);
        if (py1 + py2 == 0) {
            if (h > scaledHeight) {
                py1 = (int) (h + 10 - scaledHeight) / 2;
            } else if (h > height) {
                py1 = (int) (h + 10 - height) / 2;
            } else {
                py1 = 5;
            }
        } else {
            h -= py1 + py2;
        }

        //System.out.println("px1 = " + px1 + ",py1=" + py1);

        scale = Math.min(w / width, h / height);
        scaledWidth = width;
        scaledHeight = height;

        if (scale < 1) {
            scaledWidth = width * scale;
            scaledHeight = height * scale;
        }

        //System.out.println("   w=" + w + "," + "width=" + width + "," + "scaledWidth=" + scaledWidth);
        //System.out.println("   h=" + h + "," + "height=" + height + "," + "scaledHeight=" + scaledHeight + "\n" + px1 + "," + py1 + "|" + scale);

        int[] c = getDx(sheet, col, px1);
        int col1 = c[1];
        int dx1 = c[0];

        c = getDx(sheet, col, scaledWidth + px1);
        col2 = c[1];
        int dx2 = c[0];

        //System.out.println(col + "," + col2 + "," + dx1 + "," + dx2);

        c = getDy(sheet, row, py1);
        int row1 = c[1];
        int dy1 = c[0];

        c = getDy(sheet, row, scaledHeight + py1);
        row2 = c[1];
        int dy2 = c[0];

        //System.out.println(row + "," + row2 + "," + dy1 + "," + dy2);


        anchor.setDx1(dx1);
        anchor.setDy1(dy1);
        anchor.setDx2(dx2);
        anchor.setDy2(dy2);
        anchor.setCol1(col1);
        anchor.setCol2(col2);
        anchor.setRow1(row1);
        anchor.setRow2(row2);

        int r = workbook.addPicture(pictureData, Workbook.PICTURE_TYPE_JPEG) + insidePictureCount;
        drawing.createPicture(anchor, r);
    }

    public void createPicture(int col, int row,
                              int px1, int py1,
                              int width, int height,
                              byte[] pictureData,
                              int insidePictureCount) {
        //int dx1 = (int) (px1 / getColumnWidthInPixels(sheet, col) * 1024);
        //int dy1 = (int) (Math.abs(py1) / getRowHeightInPixels(sheet, row) * 254);

        int[] c = getDx(sheet, col, px1);
        int col1 = c[1];
        int dx1 = c[0];

        c = getDx(sheet, col, width + px1);
        int col2 = c[1];
        int dx2 = c[0];

        c = getDy(sheet, row, py1);
        int row1 = c[1];
        int dy1 = c[0];

        c = getDy(sheet, row, height + py1);
        int row2 = c[1];
        int dy2 = c[0];

        ClientAnchor anchor = creationHelper.createClientAnchor();
        anchor.setDx1(dx1);
        anchor.setDy1(dy1);
        anchor.setDx2(dx2);
        anchor.setDy2(dy2);
        anchor.setCol1(col1);
        anchor.setCol2(col2);
        anchor.setRow1(row1);
        anchor.setRow2(row2);
        int r = workbook.addPicture(pictureData, Workbook.PICTURE_TYPE_JPEG) + insidePictureCount;
        drawing.createPicture(anchor, r);
    }


    public static int[] getDx(Sheet sheet, int col, double offset) {
        double w = getColumnWidthInPixels(sheet, col);
        //System.out.println("    w=" + w + ",offset=" + offset + ",col=" + col);

        col++;

        int dx = 0;

        while (w < offset) {
            w += getColumnWidthInPixels(sheet, col++);
        }
        if (w > offset) {
            //calculate dx2, offset in the rightmost cell
            col--;
            double cw = getColumnWidthInPixels(sheet, col);
            double delta = w - offset;
            dx = (int) ((cw - delta) / cw * 1024);
        }
        //System.out.println("        dx=" + dx + ",col=" + col);
        return new int[]{dx, col};
    }

    public static int[] getDy(Sheet sheet, int row, double offset) {
        double h = getRowHeightInPixels(sheet, row);
        //System.out.println("    h=" + h + ",offset=" + offset + ",row=" + row);
        row++;
        int dy = 0;

        while (h < offset) {
            h += getRowHeightInPixels(sheet, row++);
        }
        if (h > offset) {
            row--;
            double ch = getRowHeightInPixels(sheet, row);
            double delta = h - offset;
            dy = (int) ((ch - delta) / ch * 256);
        }
        //System.out.println("        dy=" + dy + ",row=" + row);
        return new int[]{dy, row};
    }

    public static final int PX_ROW = 15;
    public static final float PX_DEFAULT = 32.00f;
    public static final float PX_MODIFIED = 36.56f;

    public static float getRowHeightInPixels(Sheet sheet, int row) {

        float height;
        Row row1 = sheet.getRow(row);
        if (row1 != null) height = row1.getHeight();
        else height = sheet.getDefaultRowHeight();

        return height / PX_ROW;
    }

    public static float getColumnWidthInPixels(Sheet sheet, int column) {
        return sheet.getColumnWidth(column) / (sheet.getDefaultColumnWidth() * 256 == sheet.getColumnWidth(column) ? PX_DEFAULT : PX_MODIFIED);
    }
}
