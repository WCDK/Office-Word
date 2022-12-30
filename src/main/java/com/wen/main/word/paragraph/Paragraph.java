package com.wen.main.word.paragraph;

import com.wen.main.word.core.CoreProperties;
import com.wen.main.word.core.WordItem;
import lombok.Data;


public class Paragraph implements WordItem {
    private String prefix = "w";
    private String name = "p";
    private RProperties pPr;
    private R txt = new R();
    private String imageUrl;
    private String imageBase64;
    private BookmarkStart bookmarkStart = new BookmarkStart();
    private BookmarkEnd bookmarkEnd = new BookmarkEnd();

    public Paragraph() {
        pPr = new RProperties();
        pPr.setName("pPr");
    }

    public RProperties getpPr() {
        return pPr;
    }

    public R getTxt() {
        return txt;
    }

    public String getPrefix() {
        return prefix;
    }

    public String getName() {
        return name;
    }

    public String getTextString() {
        return this.txt.getT();
    }

    public void setText(String text) {
        this.txt.setT(text);
    }

    public String getImageUrl() {
        return imageUrl;
    }

    public void setImageUrl(String imageUrl) {
        this.imageUrl = imageUrl;
    }

    public String getImageBase64() {
        return imageBase64;
    }

    public void setImageBase64(String imageBase64) {
        this.imageBase64 = imageBase64;
    }

    public String getFontColor() {
        return this.txt.rPr.fonts.color;
    }

    public void setFontColor(String color) {
        this.txt.setColor(color);
    }

    public String getBackground() {
        return this.txt.rPr.fonts.highlight;
    }

    public void setBackground(String background) {
        this.txt.rPr.fonts.highlight = background;
    }

    public String getpStyle() {
        return this.pPr.pStyle;
    }

    public void setpStyle(String pStyle) {
        this.pPr.pStyle = pStyle;
    }

    public String getBidi() {
        return this.pPr.bidi;
    }

    public void setBidi(String bidi) {
        this.pPr.bidi = bidi;
    }

    public class R{
        private String prefix="w";
        private String name = "r";
        private RProperties rPr = new RProperties();
        private String t;

        public String getColor() {
            return  this.rPr.fonts.color;
        }

        public void setColor(String color) {
            this.rPr.fonts.color = color;
        }

        public String getPrefix() {
            return prefix;
        }

        public String getName() {
            return name;
        }

        public RProperties getrPr() {
            return rPr;
        }

        public void setrPr(RProperties rPr) {
            this.rPr = rPr;
        }

        public String getT() {
            return t;
        }

        public void setT(String t) {
            this.t = t;
        }

        public CoreProperties toCoreProperties(){
            CoreProperties r = new CoreProperties(this.prefix,this.name);
            CoreProperties rpr = this.rPr.toCoreProperties();
            r.addChild(rpr);
            if(this.t != null){
                CoreProperties t = new CoreProperties("w","t");
                t.setValue(this.t);
                r.addChild(t);
            }

            return r;
        }
    }

    @Data
    public class BookmarkStart{
        String prefix = "w";
        String name = "bookmarkStart";
        String id;
        String wName="_GoBack";
    }
    @Data
    public class BookmarkEnd{
        String prefix = "w";
        String name = "bookmarkEnd";
        String id;
    }

    public CoreProperties toCoreProperties(){
        CoreProperties p = new CoreProperties(this.prefix,this.name);
        CoreProperties pPr = this.pPr.toCoreProperties();
        CoreProperties coreProperties = this.txt.toCoreProperties();
        p.addChild(pPr,coreProperties);
        if(this.bookmarkStart.id != null){
            CoreProperties start = new CoreProperties("w",this.bookmarkStart.name);
            CoreProperties end = new CoreProperties("w",this.bookmarkEnd.name);
            start.addAttribute("w:id",this.bookmarkStart.id);
            start.addAttribute("w:name",this.bookmarkStart.wName);
            end.addAttribute("w:id",this.bookmarkEnd.id);
            p.addChild(start,end);

        }

        return p;
    }

}
