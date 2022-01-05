package puhao.simple.msword.generator.parse;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;

import java.util.ArrayList;
import java.util.List;

public class ParsedTableCell implements DocElementUnit {
    private DocElementUnit parent;
    private List<ParsedXWPFParagraph> paragraphs = new ArrayList<>();

    private boolean hasEscapeChar;
    private boolean hasVariable;
    private boolean isExpressionCell = false;
    private DocElementUnit nextSibling;
    private XWPFTableCell cell;

    public ParsedTableCell(XWPFTableCell cell, ParsingContext context,  DocElementUnit parent) {

        this.parent = parent;
        this.cell = cell;
        ParsedXWPFParagraph pp = null;
        ParsedXWPFParagraph previousP = null;
        int index = 0;
        for (XWPFParagraph paragraph : cell.getParagraphs()) {
            pp = new ParsedXWPFParagraph(paragraph, context,  this);
            paragraphs.add(pp);
            if(pp.isExpression()){
                if(this.isExpressionCell == false) {
                    this.isExpressionCell = true;
                }else{
                    throw new IllegalArgumentException("It's not allowed to have more than one for-loop-expression" +
                            " in single table cell.");
                }
            }
            if(pp.isHasVariable()){
                this.hasVariable = true;
            }
            if(pp.isHasEscapeChar()){
                this.hasEscapeChar = true;
            }

            if(previousP != null){
                previousP.setNextSibling(pp);
            }

            previousP = pp;
            index ++;
        }
    }

    /**
     *
     * @return true is no escape char, no variable, not expression
     */
    public boolean isNormalTextCell(){
        return !(hasEscapeChar || hasVariable || isExpressionCell);
    }

    public boolean isHasEscapeChar() {
        return hasEscapeChar;
    }

    public boolean isHasVariable() {
        return hasVariable;
    }

    public boolean isExpressionCell() {
        return isExpressionCell;
    }

    public List<ParsedXWPFParagraph> getParagraphs() {
        return paragraphs;
    }

    public ParsedXWPFParagraph getParagraphs(int index){
        return paragraphs.get(index);
    }

    @Override
    public String toString() {
        return "ParsedTableCell{" +
                "paragraphs=" + paragraphs +
                '}';
    }

    @Override
    public DocElementType getType() {
        return DocElementType.TABLE_CELL;
    }

    @Override
    public Object getRawDocObject() {
        return this.cell;
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
        return this.paragraphs;
    }

    public void setNextSibling(DocElementUnit unit) {
        this.nextSibling = unit;
    }
}
