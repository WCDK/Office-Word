package com.wen.main.word.image;

import com.wen.main.word.core.CoreProperties;
import com.wen.main.word.core.WordItem;
import lombok.Data;

import java.util.List;
import java.util.stream.Collectors;

@Data
public class WordImage implements WordItem {
    String flg = "drawing";
    String prefix = "w";
    /**
     * inLine wp
     **/
    int distT;
    int distB;
    int distL;
    int distR;
    /**
     * extent wp
     **/
    /**
     * cx
     **/
    long width;

    /** cy **/
    long height;

    /**
     * effectExtent wp
     **/
    int effectExtent_T;
    int effectExtent_B;
    int effectExtent_L;
    int effectExtent_R;
    /**
     * docPr wp
     **/
    String id;
    String name;
    /** 图片路径 **/
    String picSrc;
    String rId;
    String target="media/";

    CNvGraphicFramePr cNvGraphicFramePr = new CNvGraphicFramePr();
    Graphic graphic = new Graphic();
    @Override
    public CoreProperties toCoreProperties() {
        CoreProperties drawing = new CoreProperties(prefix, flg);

        CoreProperties inline = new CoreProperties("wp", "inline");
        inline.addAttribute("distT",String.valueOf(distT));
        inline.addAttribute("distB",String.valueOf(distB));
        inline.addAttribute("distL",String.valueOf(distL));
        inline.addAttribute("distR",String.valueOf(distR));

        CoreProperties extent = new CoreProperties("wp", "extent");
        extent.addAttribute("cx",String.valueOf(width));
        extent.addAttribute("cy",String.valueOf(height));


        CoreProperties effectExtent = new CoreProperties("wp", "effectExtent");
        effectExtent.addAttribute("l",String.valueOf(effectExtent_L));
        effectExtent.addAttribute("t",String.valueOf(effectExtent_T));
        effectExtent.addAttribute("r",String.valueOf(effectExtent_R));
        effectExtent.addAttribute("b",String.valueOf(effectExtent_B));

        CoreProperties docPr = new CoreProperties("wp", "docPr");
        docPr.addAttribute("id",id);
        docPr.addAttribute("name",name);

        CoreProperties cnv = this.cNvGraphicFramePr.toCoreProperties();
        CoreProperties gra = this.graphic.toCoreProperties();
        drawing.addChild(inline,extent,effectExtent,docPr,cnv,gra);

        return drawing;
    }

    @Data
    public class CNvGraphicFramePr {
        String name = "cNvGraphicFramePr";
        String prefix = "wp";
        GraphicFrameLocks graphicFrameLocks;

        @Data
        public class GraphicFrameLocks {
            String name = "graphicFrameLocks";
            String prefix = "a";
            String a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            String noChangeAspect = "1";

            public CoreProperties toCoreProperties() {
                CoreProperties graphicFrameLocks = new CoreProperties(prefix, name);
                graphicFrameLocks.addAttribute("xmlns:a", a);
                graphicFrameLocks.addAttribute("noChangeAspect", noChangeAspect);
                return graphicFrameLocks;
            }
        }

        public CoreProperties toCoreProperties() {
            CoreProperties cNvGraphicFramePr = new CoreProperties(prefix, name);
            CoreProperties coreProperties = graphicFrameLocks.toCoreProperties();
            cNvGraphicFramePr.addChild(coreProperties);
            return cNvGraphicFramePr;
        }
    }

    @Data
    public class Graphic {
        String name = "graphic";
        String prefix = "a";
        String a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        GraphicData graphicData = new GraphicData();

        public CoreProperties toCoreProperties() {
            CoreProperties graphic = new CoreProperties(prefix, name);
            graphic.addAttribute("xmlns:a", a);
            graphic.addChild(graphicData.toCoreProperties());
            return graphic;
        }

        @Data
        public class GraphicData {
            String name = "graphicData";
            String prefix = "a";
            String uri = "http://schemas.openxmlformats.org/drawingml/2006/picture";
            Pic pic = new Pic();
            public CoreProperties toCoreProperties() {
                CoreProperties graphicData = new CoreProperties(prefix, name);
                graphicData.addAttribute("uri", uri);
                graphicData.addChild(pic.toCoreProperties());
                return graphicData;
            }

            public class Pic {
                String name = "pic";
                String prefix = "pic";
                NvPicPr nvPicPr = new NvPicPr();
                BlipFill blipFill = new BlipFill();
                SpPr spPr;
                /**
                 * xmlns
                 **/
                String pic = "http://schemas.openxmlformats.org/drawingml/2006/picture";

                public CoreProperties toCoreProperties() {
                    CoreProperties pic = new CoreProperties(prefix, name);
                    pic.addAttribute("xmlns:pic", this.pic);
                    CoreProperties nvPicPr = this.nvPicPr.toCoreProperties();
                    CoreProperties blipFill = this.blipFill.toCoreProperties();
                    CoreProperties spPr = this.spPr.toCoreProperties();
                    pic.addChild(nvPicPr,blipFill,spPr);
                    return pic;
                }

                @Data
                public class NvPicPr {
                    String name = "nvPicPr";
                    String prefix = "pic";
                    CNvPr CNvPr = new CNvPr();
                    CNvPicPr cNvPicPr = new CNvPicPr();

                    public CoreProperties toCoreProperties() {
                        CoreProperties nvPicPr = new CoreProperties(prefix, name);
                        CoreProperties CNvPr = this.CNvPr.toCoreProperties();
                        CoreProperties CNvPicPr = this.cNvPicPr.toCoreProperties();
                        nvPicPr.addChild(CNvPr, CNvPicPr);
                        return nvPicPr;
                    }

                    @Data
                    public class CNvPr {
                        String title = "cNvPr";
                        String prefix = "pic";
                        String id;
                        String name;

                        public CoreProperties toCoreProperties() {
                            CoreProperties cNvPr = new CoreProperties(prefix, name);
                            cNvPr.addAttribute("id", id);
                            cNvPr.addAttribute("name", name);
                            return cNvPr;
                        }
                    }

                    @Data
                    public class CNvPicPr {
                        String name = "cNvPicPr";
                        String prefix = "pic";
                        /**
                         * a:picLocks
                         **/
                        String noChangeAspect = "1";
                        String noChangeArrowheads = "1";

                        public CoreProperties toCoreProperties() {
                            CoreProperties cNvPicPr = new CoreProperties(prefix, name);
                            CoreProperties picLocks = new CoreProperties("a", "picLocks");
                            picLocks.addAttribute("noChangeAspect", noChangeAspect);
                            picLocks.addAttribute("noChangeArrowheads", noChangeArrowheads);
                            cNvPicPr.addChild(picLocks);
                            return cNvPicPr;
                        }
                    }

                }

                @Data
                public class BlipFill{
                    String name = "blipFill";
                    String prefix = "pic";
                    Blip blip = new Blip();
                    public CoreProperties toCoreProperties() {
                        CoreProperties blipFill = new CoreProperties(prefix, name);
                        CoreProperties blip = this.blip.toCoreProperties();
                        CoreProperties srcRect = new CoreProperties("a", "srcRect");

                        CoreProperties stretch = new CoreProperties("a", "stretch");
                        CoreProperties fillRect = new CoreProperties("a", "fillRect");
                        stretch.addChild(fillRect);
                        blipFill.addChild(blip,srcRect,stretch);
                        return blipFill;
                    }

                    @Data
                    public class Blip{
                        /**blip **/
                        String name = "blip";
                        String prefix = "a";
                        String embed;
                        String cstate = "print";
                        /**extlst **/
                        List<Ext> exts;
                        public CoreProperties toCoreProperties() {
                            CoreProperties blip = new CoreProperties(prefix, name);
                            blip.addAttribute("r:embed",embed);
                            blip.addAttribute("cstate",cstate);

                            CoreProperties extLst = new CoreProperties("a", "extLst");
                            List<CoreProperties> collect = exts.stream().map(Ext::toCoreProperties).collect(Collectors.toList());
                            extLst.addChild(collect);
                            return blip;
                        }

                        @Data
                        public class Ext{
                            String name = "ext";
                            String prefix = "a";
                            /** ext **/
                            String uri;
                            /**useLocalDpi **/
                            String a14="http://schemas.microsoft.com/office/drawing/2010/main";
                            String val="0";
                            public CoreProperties toCoreProperties() {
                                CoreProperties ext = new CoreProperties("a", "ext");
                                ext.addAttribute("uri",uri);

                                CoreProperties useLocalDpi = new CoreProperties("a14", "ext");
                                useLocalDpi.addAttribute("xmlns:a14",a14);
                                useLocalDpi.addAttribute("val",val);
                                ext.addChild(useLocalDpi);
                                return ext;
                            }


                        }
                    }

                }

                @Data
                public class SpPr{
                    String prefix = "pic";
                    String name = "spPr";
                    String bwMode = "auto";
                    Xfrm xfrm;
                    PrstGeom prstGeom;

                    public CoreProperties toCoreProperties() {
                        CoreProperties spPr = new CoreProperties(prefix, name);
                        spPr.addAttribute("bwMode", bwMode);
                        CoreProperties xfrm = this.xfrm.toCoreProperties();
                        CoreProperties prstGeom = this.prstGeom.toCoreProperties();
                        CoreProperties noFill = new CoreProperties("a", "noFill");
                        CoreProperties ln = new CoreProperties("a", "ln");
                        ln.addChild(noFill);

                        spPr.addChild(xfrm,prstGeom,noFill,ln);
                        return spPr;
                    }

                    @Data
                    public class Xfrm{
                        String prefix = "a";
                        String name = "xfrm";
                        /** off**/
                        int x;
                        int y;
                        /** ext**/
                        int cx;
                        int cy;

                        public CoreProperties toCoreProperties() {
                            CoreProperties xfrm = new CoreProperties(prefix, name);

                            CoreProperties off = new CoreProperties(prefix, "off");
                            off.addAttribute("x",String.valueOf(x));
                            off.addAttribute("y",String.valueOf(y));

                            CoreProperties ext = new CoreProperties(prefix, "ext");
                            ext.addAttribute("cx",String.valueOf(cx));
                            ext.addAttribute("cy",String.valueOf(cy));

                            xfrm.addChild(off,ext);

                            return xfrm;
                        }

                    }

                    @Data
                    public class PrstGeom{

                       String prst = "rect";

                       public CoreProperties toCoreProperties(){
                           CoreProperties prstGeom = new CoreProperties("a", "prstGeom");
                           prstGeom.addAttribute("prst",prst);
                           CoreProperties avLst = new CoreProperties("a", "avLst");
                           prstGeom.addChild(avLst);
                           return prstGeom;
                       }
                    }

                }

            }
        }
    }

    public void setId(String id){
        this.id = id;
        this.graphic.graphicData.pic.nvPicPr.getCNvPr().setId(id);
    }
    public void setName(String name){
        this.name = name;
        this.graphic.graphicData.pic.nvPicPr.getCNvPr().setName(name);
    }

    public void setrId(String rid){
        this.rId = rid;
        this.graphic.graphicData.pic.blipFill.blip.embed=rid;
    }
}
