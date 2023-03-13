package com.wen.main.word.image;

import com.wen.main.word.core.CoreProperties;
import com.wen.main.word.core.WordItem;
import lombok.AllArgsConstructor;
import lombok.Data;


@AllArgsConstructor
@Data
public class WordImage implements WordItem {
    /**
     * 距离
     **/
    int dist_l;
    int dist_t;
    int dist_b;
    int dist_r;
    /**
     * 范围
     **/
    int extent_cx;
    int extent_cy;
    int extent_l;
    int extent_t;
    int extent_r;
    int extent_b;
    int off_x;
    int off_y;
    String name;
    String id;
    String descr;
    String xmlns_a = "http://schemas.openxmlformats.org/drawingml/2006/main";
    int noChangeAspect = 1;
    String uri = "http://schemas.openxmlformats.org/drawingml/2006/picture";
    String pic = "http://schemas.openxmlformats.org/drawingml/2006/picture";
    int noChangeArrowheads = 1;
    String embed;
    String cstate = "print";
    String ext_uri = "{28A0092B-C50C-407E-A947-70E740481C1C}";
    String a14 = "http://schemas.microsoft.com/office/drawing/2010/main";
    String a14_val = "0";

    String target;
    String picSrc;
    String prst = "rect";

    public WordImage(String src){
        this.picSrc = src;
    }

    @Override
    public CoreProperties toCoreProperties() {
        CoreProperties drawing = new CoreProperties("w", "drawing");
        drawing.addChild(buildInline());
        return drawing;
    }

    private CoreProperties buildInline() {
        CoreProperties inline = new CoreProperties("wp", "inline");
        inline.addAttribute("distL", this.dist_l + "");
        inline.addAttribute("distT", this.dist_t + "");
        inline.addAttribute("distR", this.dist_r + "");
        inline.addAttribute("distB", this.dist_b + "");

        CoreProperties extent = new CoreProperties("wp", "extent");
        extent.addAttribute("cx", this.extent_cx + "");
        extent.addAttribute("cy", this.extent_cy + "");

        CoreProperties effectExtent = new CoreProperties("wp", "effectExtent");
        effectExtent.addAttribute("l", this.extent_l + "");
        effectExtent.addAttribute("t", this.extent_t + "");
        effectExtent.addAttribute("r", this.extent_r + "");
        effectExtent.addAttribute("b", this.extent_b + "");

        CoreProperties docPr = new CoreProperties("wp", "docPr");
        docPr.addAttribute("id", this.id);
        docPr.addAttribute("name", this.name);
        docPr.addAttribute("descr", this.descr);

        CoreProperties cNvGraphicFramePr = new CoreProperties("wp", "cNvGraphicFramePr");
        CoreProperties graphicFrameLocks = new CoreProperties("a", "graphicFrameLocks");
        graphicFrameLocks.addAttribute("xmlns:a", this.xmlns_a);
        graphicFrameLocks.addAttribute("noChangeAspect", this.noChangeAspect + "");
        cNvGraphicFramePr.addChild(graphicFrameLocks);
        CoreProperties graphic = buildGraphic();

        inline.addChild(extent, effectExtent, docPr, cNvGraphicFramePr, graphic);
        return inline;
    }

    private CoreProperties buildGraphic() {
        CoreProperties graphic = new CoreProperties("wp", "graphic").addAttribute("xmlns:a", this.xmlns_a);
        CoreProperties graphicData = new CoreProperties("a", "graphicData").addAttribute("uri", this.uri);
        CoreProperties pic = new CoreProperties("pic", "pic").addAttribute("xmlns:pic", this.pic);

        CoreProperties nvPicPr = new CoreProperties("pic", "nvPicPr");
        CoreProperties cNvPr = new CoreProperties("pic", "cNvPr");
        cNvPr.addAttribute("id",this.id);
        cNvPr.addAttribute("name","Picture "+id);
        cNvPr.addAttribute("descr",this.descr);

        CoreProperties cNvPicPr = new CoreProperties("pic", "cNvPicPr");
        CoreProperties picLocks = new CoreProperties("a", "picLocks").addAttribute("noChangeAspect",this.noChangeAspect+"");
        picLocks.addAttribute("noChangeArrowheads","1");
        cNvPicPr.addChild(picLocks);
        nvPicPr.addChild(cNvPr,cNvPicPr);

        CoreProperties blipFill = new CoreProperties("pic", "blipFill");
        CoreProperties blip = new CoreProperties("a", "blip").addAttribute("r:embed",this.embed);
        blip.addAttribute("cstate","print");
        CoreProperties extLst = new CoreProperties("a", "extLst");
        CoreProperties ext1 = new CoreProperties("a", "ext");
        ext1.addAttribute("uri",this.ext_uri);
        CoreProperties useLocalDpi = new CoreProperties("a14", "useLocalDpi");
        useLocalDpi.addAttribute("xmlns:a14","http://schemas.microsoft.com/office/drawing/2010/main");
        useLocalDpi.addAttribute("val","0");
        ext1.addChild(useLocalDpi);
        extLst.addChild(ext1);
        blip.addChild(extLst);

        CoreProperties srcRect = new CoreProperties("a", "srcRect");
        CoreProperties stretch = new CoreProperties("a", "stretch");
        CoreProperties fillRect = new CoreProperties("a", "fillRect");
        stretch.addChild(fillRect);

        blipFill.addChild(blip,srcRect,stretch);

        CoreProperties spPr = new CoreProperties("pic", "spPr");
        CoreProperties xfrm = new CoreProperties("a", "xfrm");
        CoreProperties off = new CoreProperties("a", "off");
        off.addAttribute("x",this.off_x+"");
        off.addAttribute("y",this.off_y+"");
        CoreProperties ext = new CoreProperties("a", "ext");
        ext.addAttribute("cx",this.extent_cx+"");
        ext.addAttribute("cy",this.extent_cy+"");
        xfrm.addChild(off,ext);
        CoreProperties prstGeom = new CoreProperties("a", "prstGeom").addAttribute("prst",this.prst);
        CoreProperties avLst = new CoreProperties("a", "avLst");
        prstGeom.addChild(avLst);
        spPr.addChild(xfrm,prstGeom);

        pic.addChild(nvPicPr,blipFill,spPr);
        graphicData.addChild(pic);
        graphic.addChild(graphicData);
        return graphic;
    }
}
