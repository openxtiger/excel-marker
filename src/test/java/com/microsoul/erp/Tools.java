package com.microsoul.erp;

import com.microsoul.erp.poi.POIParser;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

/**
 * @author tiger (tiger@microsoul.com) 2020/9/29
 */
public class Tools {
    public static void export(String outDir, String fileName,
                              Object[] o,
                              String tpl) {
        File f = new File(tpl + ".xls");
        try {
            POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(f));
            HSSFWorkbook wb = new HSSFWorkbook(fs);
            POIParser parser;
            int insidePictureCount = wb.getAllPictures().size();
            File pf = f.getParentFile();
            for (int i = o.length - 1; i >= 0; i--) {
                parser = new POIParser();
                if (o[i] == null) {
                    wb.removeSheetAt(i);
                    continue;
                }
                parser.parse(wb, o[i], i);

                parser.process(o[i], pf, insidePictureCount);
            }


            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

            for (int sheetNum = 0; sheetNum < wb.getNumberOfSheets(); sheetNum++) {
                Sheet sheet = wb.getSheetAt(sheetNum);

                for (Row r : sheet) {
                    for (Cell c : r) {
                        if (c.getCellType() == CellType.FORMULA) {
                            c.setCellFormula(c.getCellFormula());
                            evaluator.evaluateFormulaCell(c);
                        }
                    }
                }
            }
            File file = new File(outDir);
            if (!file.exists()) {
                file.mkdirs();
            }
            FileOutputStream out = new FileOutputStream(outDir + "/" + fileName);
            wb.write(out);
            out.flush();
            out.close();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            Runtime.getRuntime().gc();
        }


    }
}
