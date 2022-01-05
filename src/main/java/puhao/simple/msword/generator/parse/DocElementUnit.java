package puhao.simple.msword.generator.parse;

import java.util.List;

public interface DocElementUnit {
    DocElementType getType();
    Object getRawDocObject();


    DocElementUnit parent();


    boolean hasNextSibling();
    DocElementUnit nextSibling();

    boolean isContainerUnit();
    List<? extends DocElementUnit> getChildUnit();
}
