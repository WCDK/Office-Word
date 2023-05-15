package com.wcdk.main.core;

import lombok.Data;
import lombok.NoArgsConstructor;
import org.dom4j.Element;
import org.dom4j.Node;

import java.util.List;


@Data
@NoArgsConstructor
public class App implements OfficeItem {
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
        properties.addNameSpace("","http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");
        properties.addNameSpace("vt","http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
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
        List<Node> content = element.content();
        this.template = element.element("Template").getStringValue();
        this.pages = Integer.parseInt(element.element("Pages").getStringValue());
        this.words = Integer.parseInt(element.element("Words").getStringValue());
        this.characters = Integer.parseInt(element.element("Characters").getStringValue());
        this.lines = Integer.parseInt(element.element("Lines").getStringValue());
        this.paragraphs = Integer.parseInt(element.element("Paragraphs").getStringValue());
        this.totaltime = Integer.parseInt(element.element("TotalTime").getStringValue());
        this.scalecrop = Boolean.getBoolean(element.element("ScaleCrop").getStringValue());
        this.linksuptodate = Boolean.getBoolean(element.element("LinksUpToDate").getStringValue());
        this.characterswithspaces = Integer.parseInt(element.element("CharactersWithSpaces").getStringValue());
        this.docsecurity = Integer.parseInt(element.element("DocSecurity").getStringValue());
    }
}
