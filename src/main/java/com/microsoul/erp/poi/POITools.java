package com.microsoul.erp.poi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * 微妞分布式平台-公用工具包
 *
 * @author 广州加叁信息科技有限公司 (tiger@microsoul.com)
 * @version V1.0.0
 */
public class POITools {

    public static int getRow(int index, int cols) {
        return (index - 1) / cols + 1;
    }

    public static int getCol(int index, int cols) {
        return index % cols == 0 ? cols : index % cols;
    }

    public static void resetMergeds(int rowStart,
                                    int dataRows, int realDataRows,
                                    Sheet sheet) {
        CellRangeAddress region;
        Row row;
        Row targetRow;
        Cell sourceCell, targetCell;
        for (int i = sheet.getNumMergedRegions() - 1; i >= 0; i--) {
            region = sheet.getMergedRegion(i);

            if ((region.getFirstRow() == rowStart) &&
                    region.getLastRow() == rowStart + dataRows) {

                //
               /* System.out.println(
                        i + "====>" + region.getFirstRow() + "," + region.getLastRow() + "====" +
                                region.getFirstColumn() + "," + region.getLastColumn() + "===="
                                + rowStart + "," + (rowStart + dataRows));*/

                sheet.addMergedRegion(new CellRangeAddress(
                        region.getFirstRow(),
                        region.getFirstRow() + realDataRows,
                        region.getFirstColumn(),
                        region.getLastColumn()
                ));

                sheet.removeMergedRegion(i);
                row = sheet.getRow(rowStart + 1);
                if (row != null) {
                    for (int k = rowStart + dataRows + 1; k < rowStart + realDataRows; k++) {
                        targetRow = sheet.getRow(k);
                        setSytle(row, targetRow, region.getFirstColumn(), region.getLastColumn());
                    }
                }
                row = sheet.getRow(rowStart + dataRows);
                if (row != null) {
                    targetRow = sheet.getRow(rowStart + realDataRows);
                    setSytle(row, targetRow, region.getFirstColumn(), region.getLastColumn());
                }

            }
        }

    }

    private static void setSytle(Row row, Row targetRow, int colStart, int colEnd) {
        Cell sourceCell, targetCell;
        for (int j = colStart; j <= colEnd; j++) {
            sourceCell = row.getCell(j);
            targetCell = targetRow.getCell(j);
            if (targetCell == null) {
                targetCell = targetRow.createCell(j);
            }
            targetCell.setCellStyle(sourceCell.getCellStyle());
        }
    }


    public static void copyMergeds(int rowStart,
                                   int dataRows, int offset, int count,
                                   Sheet sheet) {
        CellRangeAddress region;
        for (int i = 0, l = sheet.getNumMergedRegions(); i < l; i++) {
            region = sheet.getMergedRegion(i);

            int n;

            if ((region.getFirstRow() >= rowStart) &&
                    region.getLastRow() <= rowStart + dataRows) {
                for (int j = 1; j <= count; j++) {
                    n = offset + dataRows * (j - 1);

                    CellRangeAddress r = region.copy();

                    r.setFirstRow(region.getFirstRow() + n);
                    r.setLastRow(region.getLastRow() + n);
                    r.setFirstColumn(region.getFirstColumn());
                    r.setLastColumn(region.getLastColumn());

                    sheet.addMergedRegion(r);
                }


            }
        }
    }


    public static void copyStyle(int rowStart, int colStart, int colEnd, int dataRows, int offset, int count,
                                 Sheet sheet) {
        Row r;

        Row targetRow;
        Cell sourceCell, targetCell;
        for (int c = 1; c <= count; c++) {
            //System.out.println("--------------------");
            for (int i = rowStart; i < rowStart + dataRows; i++) {
                r = sheet.getRow(i);

                if (r == null) continue;
                targetRow = sheet.getRow(i + offset + dataRows * (c - 1));

                if (targetRow == null)
                    targetRow = sheet.createRow(i + offset + dataRows * (c - 1));


                targetRow.setHeight(r.getHeight());
                for (int j = colStart; j <= colEnd; j++) {
                    sourceCell = r.getCell(j);
                    if (sourceCell == null) {
                        continue;
                    }
                    //System.out.println((i + 1) + "." + (i + offset + dataRows * (c - 1) + 1) + "===" + (j+1));
                    targetCell = targetRow.getCell(j);
                    if (targetCell == null)
                        targetCell = targetRow.createCell(j);

                    targetCell.setCellStyle(sourceCell.getCellStyle());
                }
            }
        }
    }


    public static Cell getCell(Sheet sheet, int moveRows, Cell targetCell) {
        return getCell(sheet, moveRows, targetCell.getColumnIndex(), targetCell);
    }

    public static Cell getCell(Sheet sheet, int moveRows, int col, Cell targetCell) {
        Row row = sheet.getRow(targetCell.getRowIndex() + moveRows);
        if (row == null) {
            row = sheet.createRow(targetCell.getRowIndex() + moveRows);
        }
        Cell cell = row.getCell(col);
        if (cell == null) {
            cell = row.createCell(col);
        }
        return cell;
    }


}
