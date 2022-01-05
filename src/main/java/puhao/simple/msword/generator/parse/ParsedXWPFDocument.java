package puhao.simple.msword.generator.parse;

import org.apache.poi.xwpf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.List;

public class ParsedXWPFDocument implements DocElementUnit{
    private static final Logger logger = LoggerFactory.getLogger("MSWordExporter");
    private int deepth = 0;
    private XWPFDocument doc;
    List<DocElementUnit> childList = new ArrayList<>();
    private List<ParsedXWPFHeader> parsedXWPFHeaders = new ArrayList<>();
    private List<ParsedXWPFFooter> parsedXWPFFooters = new ArrayList<>();
    public ParsedXWPFDocument(XWPFDocument doc){
        this.doc = doc;



        ParsingContext context = new ParsingContext();


        if(doc.getHeaderList() != null){
            for(XWPFHeader h : doc.getHeaderList()){
                ParsedXWPFHeader ph = new ParsedXWPFHeader(h, context, this);
                this.parsedXWPFHeaders.add(ph);
            }
        }
        if(doc.getFooterList() != null){
            for(XWPFFooter f : doc.getFooterList()){
                ParsedXWPFFooter pf = new ParsedXWPFFooter(f, context, this);
                this.parsedXWPFFooters.add(pf);
            }
        }


        BodyElementType type = null;
        DocElementUnit elementUnit = null;
        DocElementUnit previousUnit = null;
        int bodyIndex = 0;
        for(IBodyElement bodyElement : doc.getBodyElements()){
            type = bodyElement.getElementType();

            if(type == BodyElementType.CONTENTCONTROL){
                logger.warn("BodyElementType.CONTENTCONTROL is not supported yet");
                elementUnit = new NoParsedIBodyElement(bodyElement,  this);
            }else if(type == BodyElementType.PARAGRAPH){
                elementUnit = new ParsedXWPFParagraph((XWPFParagraph) bodyElement, context,  this);
            }else if(type == BodyElementType.TABLE){
                elementUnit = new ParsedTable((XWPFTable) bodyElement, context,  this);
            }else{
                throw new IllegalArgumentException("Unknown body element type: " + type + ". New version of MS Word? Or new version of Apache POI");
            }

            if(previousUnit != null){
                if(previousUnit instanceof ParsedTable){
                    ((ParsedTable)previousUnit).setNextSibling(elementUnit);
                }else if(previousUnit instanceof ParsedXWPFParagraph){
                    ((ParsedXWPFParagraph)previousUnit).setNextSibling(elementUnit);
                }else if(previousUnit instanceof NoParsedIBodyElement) {
                    ((NoParsedIBodyElement)previousUnit).setNextSibling(elementUnit);
                }
            }

            previousUnit = elementUnit;

            childList.add(elementUnit);
            bodyIndex++;
        }

        if(!context.isExpressCountBalanced()){
            throw new IllegalArgumentException("Unbalanced expression. Every one expression (start with ${{) must has a corresponding ${{end}}");
        }
    }

    public List<ParsedXWPFHeader> getParsedXWPFHeaders() {
        return parsedXWPFHeaders;
    }

    public List<ParsedXWPFFooter> getParsedXWPFFooters() {
        return parsedXWPFFooters;
    }

    @Override
    public DocElementType getType() {
        return DocElementType.ROOT_DOC;
    }

    @Override
    public Object getRawDocObject() {
        return doc;
    }

    @Override
    public DocElementUnit parent() {
        return null;
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
    public List<DocElementUnit> getChildUnit() {
        return this.childList;
    }

    @Override
    public String toString() {
        return "ParsedXWPFDocument{" +
                "childList=" + childList +
                '}';
    }
}
