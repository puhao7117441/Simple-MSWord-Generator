package puhao.simple.msword.generator.parse;

import org.apache.poi.xwpf.usermodel.IBodyElement;

import java.util.List;

import static puhao.simple.msword.generator.parse.DocElementType.OTHER;

public class NoParsedIBodyElement implements DocElementUnit {
    private IBodyElement element;
    private DocElementUnit nextSlibing;
    private DocElementUnit parent;

    public NoParsedIBodyElement(IBodyElement element, DocElementUnit parent){
        this.element = element;
        this.parent = parent;
    }

    public void setNextSibling(DocElementUnit unit) {
        this.nextSlibing = unit;
    }


    @Override
    public String toString() {
        return "NoParsedIBodyElement{}";
    }


    @Override
    public DocElementType getType() {
        return OTHER;
    }

    @Override
    public Object getRawDocObject() {
        return element;
    }

    @Override
    public DocElementUnit parent() {
        return parent;
    }

    @Override
    public boolean hasNextSibling() {
        return nextSlibing != null;
    }

    @Override
    public DocElementUnit nextSibling() {
        return nextSlibing;
    }

    @Override
    public boolean isContainerUnit() {
        return false;
    }

    @Override
    public List<DocElementUnit> getChildUnit() {
        return null;
    }
}
