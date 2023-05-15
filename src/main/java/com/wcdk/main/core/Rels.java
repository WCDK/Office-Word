package com.wcdk.main.core;

import lombok.Data;
import org.dom4j.Attribute;
import org.dom4j.Element;

import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

@Data
public class Rels implements OfficeItem {
     List<Node> elationships = new ArrayList<Node>();
    @Override
    public CoreProperties toCoreProperties() {
        if(elationships == null ){
            return null;
        }
        CoreProperties relationships = new CoreProperties(null,"Relationships");
        relationships.addNameSpace("","http://schemas.openxmlformats.org/package/2006/relationships");
        List<CoreProperties> collect =elationships.stream().map(e -> e.toCoreProperties()).collect(Collectors.toList());
        relationships.addChild(collect);
        return relationships;
    }
    public Rels(){}

    public Rels create(){
        Node node1 = new Node("http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties","docProps/app.xml","rId3");
        Node node2 = new Node("http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties","docProps/core.xml","rId2");
        Node node4 = new Node("http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument","word/document.xml","rId1");
        Node node3 = new Node("http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties","docProps/custom.xml","rId4");
        elationships = new ArrayList<>();
        elationships.add(node1);
        elationships.add(node2);
        elationships.add(node3);
        elationships.add(node4);
        return this;
    }

    public Rels(Element element){
        List<Element> relationship = element.elements("Relationship");
        elationships = new ArrayList<>();
        for (int i = 0;i < relationship.size();i++){
            Element el = relationship.get(i);
            Node relation = new Node();
            Class<? extends Node> aClass = relation.getClass();
            List<Attribute> attributes = el.attributes();
            attributes.forEach(at->{
                try {
                    Method declaredMethod = aClass.getDeclaredMethod("set" + at.getName(),String.class);
                    declaredMethod.invoke(relation,at.getStringValue());
                }catch (Exception e){
                    e.printStackTrace();
                }
            });
            elationships.add(relation);
        }
    }
    @Data
    public class Node{
        String type;
        String target;
        String id;
        public Node(String type,String target,String id){
            this.type = type;
            this.target = target;
            this.id = id;
        }
        public Node(){}
        public CoreProperties toCoreProperties() {
            CoreProperties relationship = new CoreProperties(null,"Relationship");
            relationship.addAttribute("Type",this.type);
            relationship.addAttribute("Target",this.target);
            relationship.addAttribute("Id",this.id);

            return relationship;
        }
    }
}
