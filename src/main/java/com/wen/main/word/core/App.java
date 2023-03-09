package com.wen.main.word.core;

import lombok.Data;
import lombok.NoArgsConstructor;
import org.dom4j.Element;


@Data
@NoArgsConstructor
public class App implements WordItem{
    String  template;
    int pages;
    int words;
    int characters = 4;
    int lines;
    int paragraphs;
    int totaltime = 0;
    boolean scalecrop = false;
    boolean linksuptodate = false;
    int characterswithspaces = 4;
    String application="WCDK OPEN OFFICE";
    int docsecurity;
    @Override
    public CoreProperties toCoreProperties() {
        CoreProperties properties = new CoreProperties(null,"Properties");
        properties.addAttribute("xmlns","http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");
        properties.addAttribute("xmlns:vt","http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
        CoreProperties Template = new CoreProperties(null,"Template",this.template+"");
        CoreProperties Pages = new CoreProperties(null,"Pages",this.pages+"");
        CoreProperties Words = new CoreProperties(null,"Words",this.words+"");
        CoreProperties Characters = new CoreProperties(null,"Characters",this.characters+"");
        CoreProperties Lines = new CoreProperties(null,"Lines",this.lines+"");
        CoreProperties Paragraphs = new CoreProperties(null,"Paragraphs",this.paragraphs+"");
        CoreProperties TotalTime = new CoreProperties(null,"TotalTime",this.totaltime+"");
        CoreProperties ScaleCrop = new CoreProperties(null,"ScaleCrop",this.scalecrop+"");
        CoreProperties LinksUpToDate = new CoreProperties(null,"LinksUpToDate",this.linksuptodate+"");
        CoreProperties CharactersWithSpaces = new CoreProperties(null,"CharactersWithSpaces",this.characterswithspaces+"");
        CoreProperties Application = new CoreProperties(null,"Application",this.application);
        CoreProperties DocSecurity = new CoreProperties(null,"DocSecurity",this.docsecurity+"");
        properties.addChild(Template,Pages,Words,Characters,Lines,Paragraphs,TotalTime);
        properties.addChild(ScaleCrop,LinksUpToDate,CharactersWithSpaces,Application,DocSecurity);
        return properties;
    }

    public App(Element element){
        this.template = element.selectSingleNode("Template").getStringValue();
        this.pages = Integer.parseInt(element.selectSingleNode("Pages").getStringValue());
        this.words = Integer.parseInt(element.selectSingleNode("Words").getStringValue());
        this.characters = Integer.parseInt(element.selectSingleNode("Characters").getStringValue());
        this.lines = Integer.parseInt(element.selectSingleNode("Lines").getStringValue());
        this.paragraphs = Integer.parseInt(element.selectSingleNode("Paragraphs").getStringValue());
        this.totaltime = Integer.parseInt(element.selectSingleNode("TotalTime").getStringValue());
        this.scalecrop = Boolean.getBoolean(element.selectSingleNode("ScaleCrop").getStringValue());
        this.linksuptodate = Boolean.getBoolean(element.selectSingleNode("LinksUpToDate").getStringValue());
        this.characterswithspaces = Integer.parseInt(element.selectSingleNode("CharactersWithSpaces").getStringValue());
        this.docsecurity = Integer.parseInt(element.selectSingleNode("DocSecurity").getStringValue());
    }
}
