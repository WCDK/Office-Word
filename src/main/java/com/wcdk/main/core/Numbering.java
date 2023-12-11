package com.wcdk.main.core;

import lombok.Data;
import org.dom4j.Element;

import java.util.LinkedList;
import java.util.List;

@Data
public class Numbering implements OfficeItem {
    String wpc = "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas";
    String mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    String o = "urn:schemas-microsoft-com:office:office";
    String r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    String m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    String v = "urn:schemas-microsoft-com:vml";
    String wp14 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";
    String wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    String w10 = "urn:schemas-microsoft-com:office:word";
    String w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    String w14 = "http://schemas.microsoft.com/office/word/2010/wordml";
    String wpg = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
    String wpi = "http://schemas.microsoft.com/office/word/2010/wordprocessingInk";
    String wne = "http://schemas.microsoft.com/office/word/2006/wordml";
    String wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
    String wpsCustomData = "http://www.wps.cn/officeDocument/2013/wpsCustomData";
    String mc_Ignorable = "w14 wp14";
    public List<Lvl> lvls = new LinkedList<>();
    String numId="1";
    String val = "0";
    String abstractNumId = "0";
    String nsid = "BC1EEE41";
    String multiLevelType = "multilevel";
    String tmpl = "BC1EEE41";
    Lvl lastLvl;
    boolean used = false;
    public Numbering(){}

    public Numbering(Element root){
        this.used = true;
        Element abstractNum = root.element("abstractNum");
        this.abstractNumId = abstractNum.attribute("w:abstractNumId").getStringValue();
        this.nsid = abstractNum.element("nsid").attribute("w:val").getStringValue();
        this.multiLevelType = abstractNum.element("multiLevelType").attribute("w:val").getStringValue();
        this.tmpl = abstractNum.element("tmpl").attribute("w:val").getStringValue();

        List<Element> lvlElement = abstractNum.elements("lvl");
        lvlElement.forEach(e->lvls.add(new Lvl(e)));
        Element num = abstractNum.element("num");
        this.numId = num.attribute("w:numId").getStringValue();
    }

    @Override
    public CoreProperties toCoreProperties() {
        CoreProperties numbering = new CoreProperties("w", "numbering");
        numbering.addAttribute("xmlns:wpc", this.wpc);
        numbering.addAttribute("xmlns:mc", this.mc);
        numbering.addAttribute("xmlns:o", this.o);
        numbering.addAttribute("xmlns:r", this.r);
        numbering.addAttribute("xmlns:m", this.m);
        numbering.addAttribute("xmlns:v", this.v);
        numbering.addAttribute("xmlns:wp14", this.wp14);
        numbering.addAttribute("xmlns:wp", this.wp);
        numbering.addAttribute("xmlns:w10", this.w10);
        numbering.addAttribute("xmlns:w", this.w);
        numbering.addAttribute("xmlns:w14", this.w14);
        numbering.addAttribute("xmlns:wpg", this.wpg);
        numbering.addAttribute("xmlns:wpi", this.wpi);
        numbering.addAttribute("xmlns:wne", this.wne);
        numbering.addAttribute("xmlns:wps", this.wps);
        numbering.addAttribute("xmlns:wpsCustomData", this.wpsCustomData);
        numbering.addAttribute("mc:Ignorable", this.mc_Ignorable);
        CoreProperties absTractNum = createAbsTractNum();
        CoreProperties num = createNum();
        numbering.addChild(absTractNum,num);

        return numbering;
    }

    private CoreProperties createAbsTractNum() {
        CoreProperties abstractNum = new CoreProperties("w", "abstractNum");
        abstractNum.addAttribute("w:abstractNumId",this.abstractNumId);
        CoreProperties nsid = new CoreProperties("w", "nsid").addAttribute("w:val",this.nsid);
        CoreProperties multiLevelType = new CoreProperties("w", "multiLevelType").addAttribute("w:val",this.multiLevelType);
        CoreProperties tmpl = new CoreProperties("w", "tmpl").addAttribute("w:val",this.tmpl);
        abstractNum.addChild(nsid,multiLevelType,tmpl);
        lvls.forEach(e->abstractNum.addChild(e.toCoreProperties()));
        return abstractNum;
    }

    private CoreProperties createNum(){
        CoreProperties num = new CoreProperties("w","num").addAttribute("w:numId",this.numId);
        CoreProperties abstractNumId = new CoreProperties("w","abstractNumId").addAttribute("w:val",this.val);
        num.addChild(abstractNumId);
        return num;
    }

    public void buildLvl(String pStyle){
        if(lastLvl == null){
            Lvl lvl = new Lvl();
            lastLvl = lvl;
            lvls.add(lvl);
            return;
        }
        Lvl lvl = lastLvl.getNext(pStyle);
        lastLvl = lvl;
        lvls.add(lvl);
    }

    @Data
    public class Lvl{
        int curr = 1;
        String ilvl = "0";
        String tentative = "0";
        String start = "1";
        String numFmt = "decimal";
        String pStyle = "2";
        String lvlText = "%1.";
        String lvlJc = "left";
        String left = "432";
        String hanging = "432";
        String fonts = "default";
        public CoreProperties toCoreProperties(){
            CoreProperties lvl = new CoreProperties("w","lvl");
            lvl.addAttribute("w:ilvl",ilvl);
            lvl.addAttribute("w:tentative","0");
            CoreProperties start = new CoreProperties("w","start").addAttribute("w:val",this.start);
            CoreProperties numFmt = new CoreProperties("w","numFmt").addAttribute("w:val",this.numFmt);
            CoreProperties pStyle = new CoreProperties("w","pStyle").addAttribute("w:val",this.pStyle);
            CoreProperties lvlText = new CoreProperties("w","lvlText").addAttribute("w:val",this.lvlText);
            CoreProperties lvlJc = new CoreProperties("w","lvlJc").addAttribute("w:val",this.lvlJc);

            CoreProperties pPr = new CoreProperties("w","pPr");
            CoreProperties ind = new CoreProperties("w","ind");
            ind.addAttribute("w:left",this.left);
            ind.addAttribute("w:hanging",this.hanging);
            pPr.addChild(ind);
            CoreProperties rPr = new CoreProperties("w","rPr");
            CoreProperties rFonts = new CoreProperties("w","rFonts");
            rFonts.addAttribute("w:hint",this.fonts);
            rPr.addChild(rFonts);

            lvl.addChild(start,numFmt,pStyle,lvlText,lvlJc,pPr,rPr);
            return lvl;
        }
        public Lvl(){};
        public Lvl(Element root){
            this.ilvl = root.attribute("w:ilvl").getStringValue();
            this.tentative = root.attribute("w:tentative").getStringValue();
            this.start = root.element("start").attribute("w:val").getStringValue();
            this.numFmt = root.element("numFmt").attribute("w:val").getStringValue();
            this.pStyle = root.element("pStyle").attribute("w:val").getStringValue();
            this.lvlText = root.element("lvlText").attribute("w:val").getStringValue();
            this.lvlJc = root.element("lvlJc").attribute("w:val").getStringValue();

            Element pPr = root.element("pPr");
            this.left = pPr.element("ind").attribute("w:left").getStringValue();
            this.hanging = pPr.element("ind").attribute("w:hanging").getStringValue();

            Element rPr = root.element("rPr");
            this.fonts = rPr.element("rFonts").attribute("w:hint").getStringValue();
        }

        public Lvl getNext(String pStyle){
            Lvl lvl = new Lvl();
            lvl.setCurr(curr+1);
            lvl.setPStyle(pStyle);
            lvl.setIlvl(String.valueOf(Integer.parseInt(ilvl)+1));
            lvl.setLvlText(String.format("%s%s.",lvlText,"%"+(curr+1)));
            return lvl;
        }
    }
}
