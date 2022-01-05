package puhao.simple.msword.generator.parse;

import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.util.ArrayList;
import java.util.List;

public class ParsedXWPFHeader implements DocElementUnit{
    private List<ParsedXWPFParagraph> paragraphs = new ArrayList<>();
    private XWPFHeader header;
    private DocElementUnit parent;

    public ParsedXWPFHeader(XWPFHeader header, ParsingContext context, DocElementUnit parent) {
        this.parent = parent;
        List<XWPFParagraph> paragraphs = header.getParagraphs();
        if(paragraphs != null){
            for(XWPFParagraph paragraph : paragraphs) {
                ParsedXWPFParagraph parsedP = new ParsedXWPFParagraph(paragraph, context, this);
                this.paragraphs.add(parsedP);
                if(parsedP.isExpression()){
                    throw new IllegalArgumentException("Not allowed to use for-loop-expression in document header");
                }
            }
        }
    }

    public List<ParsedXWPFParagraph> getParagraphs() {
        return paragraphs;
    }

    public ParsedXWPFParagraph getParagraph(int i) {
        return paragraphs.get(i);
    }
    @Override
    public DocElementType getType() {
        return DocElementType.HEADER_FOOTER;
    }

    @Override
    public Object getRawDocObject() {
        return header;
    }

    @Override
    public DocElementUnit parent() {
        return parent;
    }

    @Override
    public boolean hasNextSibling() {
        return false;
    }

    @Override
    public DocElementUnit nextSibling() {
        return null;
    }

    @Override
    public boolean isContainerUnit() {
        return true;
    }

    @Override
    public List<? extends DocElementUnit> getChildUnit() {
        return this.paragraphs;
    }

}
