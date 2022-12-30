package com.wen.main.word.paragraph;

import com.wen.main.word.core.CoreProperties;
import lombok.Data;

@Data
public class RProperties {
    String prefix = "w";
    String name = "rPr";
    String pStyle;
    String bidi = "0";
    String algin;

    Fonts fonts = new Fonts();
    Lang lang = new Lang();
    public class Fonts{
        String eastAsiaTheme;
        String hint = "eastAsia";

        String fontFace;
        String vertAlign;
        String color;
        int sz;
        String highlight;

        int szCs;
        String ascii;
        String cs;
        String eastAsia;
        String hAnsi;

        public void setFontFace(String fontFace) {
            this.fontFace = fontFace;
            this.ascii = fontFace;
            this.cs = fontFace;
            this.eastAsia = fontFace;
            this.hAnsi = fontFace;
        }
        public void setColor(String color) {
            this.color = color;
        }
        public String getColor() {
            return color;
        }
        public CoreProperties toCoreProperties(){
            CoreProperties rFonts = new CoreProperties("w","rFonts");
            if(this.fontFace != null && !this.fontFace.trim().equals("")){
                rFonts.addAttribute("w:ascii",this.ascii);
                rFonts.addAttribute("w:cs",this.cs);
                rFonts.addAttribute("w:eastAsia",this.eastAsia);
                rFonts.addAttribute("w:hAnsi",this.hAnsi);
            }
            if(this.eastAsiaTheme != null && !this.eastAsiaTheme.trim().equals("")){
                rFonts.addAttribute("w:eastAsiaTheme",this.eastAsiaTheme);
            }else  if(this.hint != null && !this.hint.trim().equals("")){
                rFonts.addAttribute("w:hint",this.hint);
            }
            return rFonts;
        }
    }
    @Data
    public class Lang{
        String prefix = "w";
        String name = "lang";
        String val= "en-US";
        String eastAsia= "zh-CN";
    }

    public CoreProperties toCoreProperties(){
        CoreProperties rpr = new CoreProperties(prefix,name);
        if(this.pStyle != null){
            CoreProperties pStyle = new CoreProperties("w","pStyle");
            pStyle.addAttribute("w:val",this.pStyle);
            CoreProperties bidi = new CoreProperties("w","bidi");
            bidi.addAttribute("w:val",this.bidi);
            rpr.addChild(pStyle,bidi);
            this.fonts.hint = "default";
        }
        CoreProperties fontElement = this.fonts.toCoreProperties();
        rpr.addChild(fontElement);

        if(this.fonts.color != null && !this.fonts.color.trim().equals("")){
            CoreProperties color = new CoreProperties("w","color");
            color.addAttribute("w:val",this.fonts.color);
            rpr.addChild(color);
        }
        if(this.fonts.sz > 0){
            CoreProperties sz = new CoreProperties("w","sz");
            CoreProperties szCs = new CoreProperties("w","szCs");
            sz.addAttribute("w:val",this.fonts.sz+"");
            sz.addAttribute("w:val",this.fonts.szCs+"");
            rpr.addChild(sz,szCs);
        }
        if(this.fonts.highlight != null && !this.fonts.highlight.trim().equals("")){
            CoreProperties highlight = new CoreProperties("w","highlight");
            highlight.addAttribute("w:val",this.fonts.highlight);
            rpr.addChild(highlight);
        }
        if(this.fonts.vertAlign != null && !this.fonts.vertAlign.trim().equals("")){
            CoreProperties vertAlign = new CoreProperties("w:vertAlign",this.fonts.vertAlign);
            rpr.addChild(vertAlign);
        }

        CoreProperties lang = new CoreProperties(this.lang.prefix,this.lang.name);
        lang.addAttribute("w:val",this.lang.val);
        lang.addAttribute("w:eastAsia",this.lang.eastAsia);
        rpr.addChild(lang);
        if(this.algin != null){
            CoreProperties jc = new CoreProperties("w","jc");
            jc.addAttribute("w:val",this.algin);
            rpr.addChild(jc);
        }
        return rpr;
    }
}
