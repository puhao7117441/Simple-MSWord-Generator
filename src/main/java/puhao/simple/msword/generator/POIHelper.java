package puhao.simple.msword.generator;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

public class POIHelper {

    public static void cloneDocumentStyles(XWPFDocument target, XWPFDocument source){


    }

    public static void cloneRow(XWPFTableRow targetRow, XWPFTableRow source){
        targetRow.getCtRow().setTrPr(source.getCtRow().getTrPr());
        for (int c=0; c<source.getTableCells().size(); c++) {
            //newly created row has 1 cell
            XWPFTableCell targetCell = c==0 ? targetRow.getTableCells().get(0) : targetRow.createCell();
            XWPFTableCell cell = source.getTableCells().get(c);
            targetCell.getCTTc().setTcPr(cell.getCTTc().getTcPr());

            while(targetCell.getParagraphs().size() > 0){
                targetCell.removeParagraph(0);
            }


            for (int p = 0; p < cell.getParagraphs().size(); p++) {
                XWPFParagraph paragraph = cell.getParagraphArray(p);
                XWPFParagraph targetParagraph = targetCell.addParagraph();

                POIHelper.cloneParagraph(targetParagraph, paragraph);
            }
        }
    }


    /**
     * I don't known what does this method did. Apache POI is magic box
     * @param source
     * @param target
     */
    public static void cloneTable(XWPFTable target, XWPFTable source) {
        target.getCTTbl().setTblPr(source.getCTTbl().getTblPr());
        target.getCTTbl().setTblGrid(source.getCTTbl().getTblGrid());
        for (int r = 0; r<source.getRows().size(); r++) {
            XWPFTableRow targetRow = target.createRow();
            XWPFTableRow row = source.getRows().get(r);
            cloneRow(targetRow, row);
        }
        //newly created table has one row by default. we need to remove the default row.
        target.removeRow(0);
    }

    public static void cloneParagraph(XWPFParagraph clone, XWPFParagraph source) {
        CTPPr pPr = clone.getCTP().isSetPPr() ? clone.getCTP().getPPr() : clone.getCTP().addNewPPr();
        pPr.set(source.getCTP().getPPr());
        for (XWPFRun r : source.getRuns()) {
            XWPFRun nr = clone.createRun();
            cloneRun(nr, r);
        }
    }

    public static void cloneRun(XWPFRun clone, XWPFRun source) {
        CTRPr rPr = clone.getCTR().isSetRPr() ? clone.getCTR().getRPr() : clone.getCTR().addNewRPr();
        rPr.set(source.getCTR().getRPr());
        clone.setText(source.getText(0));
    }

}
