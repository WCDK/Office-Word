package com.wcdk.main.core;

import lombok.Data;
import lombok.NoArgsConstructor;
import org.dom4j.Element;


@Data
@NoArgsConstructor
public class Core implements OfficeItem {
    String cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    String dc = "http://purl.org/dc/elements/1.1/";
    String dcterms="http://purl.org/dc/terms/";
    String dcmitype = "http://purl.org/dc/dcmitype/";
    String xsi = "http://www.w3.org/2001/XMLSchema-instance";
    String created = CoreProperties.getTimeString();
    String createdType="dcterms:W3CDTF";
    String creator = "WCDK";
    String lastModifiedBy = "WCDK";
    String modified = CoreProperties.getTimeString();
    String modifiedType="dcterms:W3CDTF";
    String revision = "1";
    @Override
    public CoreProperties toCoreProperties() {
        CoreProperties coreProperties = new CoreProperties("cp","coreProperties");
        coreProperties.addNameSpace("cp",this.cp);
        coreProperties.addNameSpace("dc",this.dc);
        coreProperties.addNameSpace("dcterms",this.dcterms);
        coreProperties.addNameSpace("dcmitype",this.dcmitype);
        coreProperties.addNameSpace("xsi",this.xsi);

        CoreProperties created = new CoreProperties("dcterms","created",this.created);
        created.addAttribute("xsi:type",this.createdType);

        CoreProperties creator = new CoreProperties("dc","creator",this.creator);
        CoreProperties lastModifiedBy = new CoreProperties("cp","lastModifiedBy","WCDK");

        CoreProperties modified = new CoreProperties("dcterms","modified");
        modified.addAttribute("xsi:type",this.modifiedType);
        modified.setValue(modified.getTimeString());
        CoreProperties revision = new CoreProperties("cp","revision",this.revision);
        coreProperties.addChild(created,creator,lastModifiedBy,modified,revision);

        return coreProperties;
    }

    public Core(Element e){
        this.creator = e.selectSingleNode("dc:creator").getStringValue();
        this.lastModifiedBy = "WCDK";
        this.revision = e.selectSingleNode("cp:revision").getStringValue();

        Element e1 = (Element) e.selectSingleNode("dcterms:created");
        this.createdType = e1.attributes().get(0).getStringValue();
        this.created = e1.getStringValue();

        e1 = (Element) e.selectSingleNode("dcterms:modified");
        this.modifiedType = e1.attributes().get(0).getStringValue();
        this.modified = e1.getStringValue();
    }

    public void setLastModifiedBy(String lastModifiedBy) {
        this.lastModifiedBy = "WCDK";
    }
}
