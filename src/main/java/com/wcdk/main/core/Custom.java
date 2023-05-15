package com.wcdk.main.core;

import lombok.Data;
import org.dom4j.Element;

import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

@Data
public class Custom implements OfficeItem {
    String propertyxmlns = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
    String propertyxmlnsVt = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
    List<Property> properties = new ArrayList<>();
    @Override
    public CoreProperties toCoreProperties() {
        if(properties.size() < 1){
            return null;
        }
        CoreProperties coreProperties = new CoreProperties(null,"Properties");
        coreProperties.addNameSpace("",this.propertyxmlns);
        coreProperties.addNameSpace("vt",this.propertyxmlnsVt);
        List<CoreProperties> collect = properties.stream().map(e -> e.toCoreProperties()).collect(Collectors.toList());
        coreProperties.addChild(collect);

        return coreProperties;
    }
    public Custom(){}
    public Custom create(){
        Property property1 = new Property();
        Property property2 = new Property();

        property1.setFmtid("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}");
        property1.setName("KSOProductBuildVer");
        property1.setPid("1");
        property1.setLpwstr("2052-11.1.0.12132");
        property2.setFmtid("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}");
        property2.setName("ICV");
        property2.setPid("2");
        property2.setLpwstr("6E3A6E7992FF46D8A6CD25090CB6CA05");
        this.properties.add(property1);
        this.properties.add(property2);
        return this;
    }
    public Custom(Element element){

        this.propertyxmlns = element.getNamespaceForPrefix("").getStringValue();
        this.propertyxmlns = element.getNamespaceForPrefix("vt").getStringValue();
        List<Element> property = element.elements("property");
        for(Element e : property){
            Property property1 = new Property();
            property1.setFmtid(e.attribute("fmtid").getStringValue());
            property1.setName(e.attribute("name").getStringValue());
            property1.setPid(e.attribute("pid").getStringValue());
            property1.setLpwstr(e.selectSingleNode("vt:lpwstr").getStringValue());
            properties.add(property1);
        }

    }

    @Data
    public class Property{
        String fmtid;
        String pid;
        String name;
        String lpwstr;

        public CoreProperties toCoreProperties() {
            CoreProperties coreProperties = new CoreProperties(null,"property");
            coreProperties.addAttribute("fmtid",this.fmtid);
            coreProperties.addAttribute("pid",this.pid);
            coreProperties.addAttribute("name",this.name);
            CoreProperties lpwstr = new CoreProperties("vt","lpwstr",this.lpwstr);
            coreProperties.addChild(lpwstr);
            return coreProperties;
        }
    }
}
