package puhao.simple.msword.generator.parse;

import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.util.ArrayList;
import java.util.List;

public class ParsedTableRow implements DocElementUnit, ExpressionContent {
    private DocElementUnit parent;
    private List<ParsedTableCell> cells = new ArrayList<>();
    private boolean isExpressionRow = false;

    private DocElementUnit nextSibling;
    private XWPFTableRow row;

    public ParsedTableRow(XWPFTableRow row, ParsingContext context, DocElementUnit parent) {
        this.parent = parent;
        this.row = row;
        ParsedTableCell c = null;
        ParsedTableCell previousC = null;
        int index = 0;
        for (XWPFTableCell cell : row.getTableCells()) {
            c = new ParsedTableCell(cell, context, this);
            cells.add(c);
            if(c.isExpressionCell()){
                if(this.isExpressionRow == false) {
                    this.isExpressionRow = true;
                }else{
                    throw new IllegalArgumentException("It's not allowed to have more than one for-loop-expression" +
                            " in single table row.");
                }
            }

            if(previousC != null){
                previousC.setNextSibling(c);
            }

            previousC = c;
            index++;
        }
    }


    public boolean isExpression() {
        return isExpressionRow;
    }


    public List<ParsedTableCell> getCells() {
        return cells;
    }
    public ParsedTableCell getCell(int index){
        return cells.get(index);
    }

    public ForLoopExpression getForLoopExpression(){
        for(ParsedTableCell cell : this.cells){
            if(cell.isExpressionCell()){
                for(ParsedXWPFParagraph p : cell.getParagraphs()){
                    if(p.isExpression()){
                        return p.getForLoopExpression();
                    }
                }
            }
        }
        return null;
    }

    @Override
    public String toString() {
        return "ParsedTableRow{" +
                "cells=" + cells +
                '}';
    }

    @Override
    public DocElementType getType() {
        return DocElementType.TABLE_ROW;
    }

    @Override
    public Object getRawDocObject() {
        return this.row;
    }


    @Override
    public DocElementUnit parent() {
        return parent;
    }


    @Override
    public boolean hasNextSibling() {
        return this.nextSibling != null;
    }

    @Override
    public DocElementUnit nextSibling() {
        return nextSibling;
    }

    @Override
    public boolean isContainerUnit() {
        return true;
    }

    @Override
    public List<? extends DocElementUnit> getChildUnit() {
        return this.cells;
    }

    public void setNextSibling(DocElementUnit unit) {
        this.nextSibling = unit;
    }
}
