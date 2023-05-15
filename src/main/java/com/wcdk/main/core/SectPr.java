package com.wcdk.main.core;

import com.wcdk.main.core.eunm.PaperType;
import lombok.Data;
import org.dom4j.Element;

@Data
public class SectPr implements OfficeItem {
    int margin_top;
    int margin_right;
    int margin_bottom;
    int margin_left;
    int margin_header;
    int margin_footer;
    int margin_gutter;
    int w;
    int h;
    int space = 425;
    String grid_type = "lines";
    int linePitch = 312;

    SectPr(PaperType paperType) {
        margin_top = paperType.getTop();
        margin_right = paperType.getRight();
        margin_bottom = paperType.getBottom();
        margin_left = paperType.getLeft();
        margin_header = paperType.getHeader();
        margin_footer = paperType.getFooter();
        margin_gutter = paperType.getGutter();
        w = paperType.getW();
        h = paperType.getH();
    }

    SectPr(Element element) {
        Element ce = element.element("pgSz");
        this.w = Integer.parseInt(ce.attribute("w:w").getStringValue());
        this.h = Integer.parseInt(ce.attribute("w:h").getStringValue());

        ce = element.element("pgMar");
        this.margin_top = Integer.parseInt(ce.attribute("w:top").getStringValue());
        this.margin_right = Integer.parseInt(ce.attribute("w:right").getStringValue());
        this.margin_bottom = Integer.parseInt(ce.attribute("w:bottom").getStringValue());
        this.margin_left= Integer.parseInt(ce.attribute("w:left").getStringValue());
        this.margin_header= Integer.parseInt(ce.attribute("w:header").getStringValue());
        this.margin_footer= Integer.parseInt(ce.attribute("w:footer").getStringValue());
        this.margin_gutter= Integer.parseInt(ce.attribute("w:gutter").getStringValue());

        ce = element.element("cols");
        this.space = Integer.parseInt(ce.attribute("w:space").getStringValue());

        ce = element.element("docGrid");
        this.grid_type = ce.attribute("w:type").getStringValue();
        this.linePitch = Integer.parseInt(ce.attribute("w:linePitch").getStringValue());

    }

    @Override
    public CoreProperties toCoreProperties() {
        CoreProperties sectPr = new CoreProperties("w","sectPr");

        CoreProperties pgSz = new CoreProperties("w","pgSz");
        pgSz.addAttribute("w:w",this.w+"");
        pgSz.addAttribute("w:h",this.h+"");

        CoreProperties pgMar = new CoreProperties("w","pgMar");
        pgMar.addAttribute("w:top",this.margin_top+"");
        pgMar.addAttribute("w:right",this.margin_right+"");
        pgMar.addAttribute("w:bottom",this.margin_bottom+"");
        pgMar.addAttribute("w:left",this.margin_left+"");
        pgMar.addAttribute("w:header",this.margin_header+"");
        pgMar.addAttribute("w:footer",this.margin_footer+"");
        pgMar.addAttribute("w:gutter",this.margin_gutter+"");

        CoreProperties cols = new CoreProperties("w","cols");
        cols.addAttribute("w:space",this.space+"");

        CoreProperties docGrid = new CoreProperties("w","docGrid");
        docGrid.addAttribute("w:type",this.grid_type);
        docGrid.addAttribute("w:linePitch",this.linePitch+"");

        sectPr.addChild(pgSz,pgMar,cols,docGrid);
        return sectPr;
    }
}
