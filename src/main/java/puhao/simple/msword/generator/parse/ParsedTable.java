package puhao.simple.msword.generator.parse;

import org.apache.poi.xwpf.usermodel.*;

import java.util.ArrayList;
import java.util.List;

public class ParsedTable implements DocElementUnit {

    List<ParsedTableRow> rows = new ArrayList<>();
    XWPFTable table = null;
    private boolean hasExpression = false;
    private DocElementUnit nextSibling;
    private DocElementUnit parent;

    public ParsedTable(XWPFTable table, ParsingContext context, DocElementUnit parent) {
        this.table = table;
        this.parent = parent;

        ParsedTableRow pr = null;
        ParsedTableRow previousPr = null;
        int index = 0;
        for(XWPFTableRow row : table.getRows()){
            try {
                pr = new ParsedTableRow(row, context, this);
            }catch (IllegalArgumentException e){
                throw new IllegalArgumentException("Error at row with index (0 based): " + index, e);
            }
            rows.add(pr);

            if(pr.isExpression()){
                this.hasExpression = true;
            }

            if(previousPr != null){
                previousPr.setNextSibling(pr);
            }

            previousPr = pr;
            index++;
        }
    }

    public boolean isHasExpression() {
        return hasExpression;
    }


    public List<ParsedTableRow> getRows() {
        return rows;
    }
    public ParsedTableRow getRow(int index) {
        return rows.get(index);
    }

    @Override
    public String toString() {
        return "ParsedTable{" +
                "rows=" + rows +
                '}';
    }

    @Override
    public DocElementType getType() {
        return DocElementType.TABLE;
    }

    @Override
    public Object getRawDocObject() {
        return this.table;
    }


    @Override
    public DocElementUnit parent() {
        return parent;
    }


    @Override
    public boolean hasNextSibling() {
        return nextSibling != null;
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
        return rows;
    }

    public void setNextSibling(DocElementUnit nextSibling) {
        this.nextSibling = nextSibling;
    }
}

