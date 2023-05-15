package com.wcdk.main.core;

import com.wcdk.main.core.eunm.RelationshipType;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.dom4j.Element;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.stream.Collectors;


@Data
public class DocumentRels implements OfficeItem {
    List<Node> relationships = new ArrayList<>();
    String xmlns = RelationshipType.base.getValue();
    boolean excel = false;
    Integer sheets;

    @Override
    public CoreProperties toCoreProperties() {
        CoreProperties relation = new CoreProperties(null,"Relationships");
        relation.addNameSpace("",this.xmlns);
        relation.addChild(relationships.stream().map(e->e.toCoreProperties()).collect(Collectors.toList()));
        return relation;
    }
    public DocumentRels(){}
    public DocumentRels(Element element){
        this.xmlns = element.getNamespaceForPrefix("").getStringValue();
        List<Element> relationship = element.elements("Relationship");
        for(Element  e : relationship){
            RelationshipType type = RelationshipType.getTyppe(e.attribute("Type").getStringValue());
            String target = e.attribute("Target").getStringValue();
            String id = e.attribute("Id").getStringValue();

            Node no = new Node(type,target,id);
            relationships.add(no);
        }
    }
    public DocumentRels create(){
        Node style = new Node(RelationshipType.styles,"styles.xml","rId"+1);
        Node settings = new Node(RelationshipType.settings,"settings.xml","rId"+2);
        Node webSettings = new Node(RelationshipType.webSettings,"webSettings.xml","rId"+3);
        Node footnotes = new Node(RelationshipType.footnotes,"footnotes.xml","rId"+4);
        Node endnotes = new Node(RelationshipType.endnotes,"endnotes.xml","rId"+5);
        Node fontTable = new Node(RelationshipType.fontTable,"fontTable.xml","rId"+6);
        Node theme = new Node(RelationshipType.theme,"theme/theme1.xml","rId"+7);
        relationships.add(style);
        relationships.add(settings);
        relationships.add(webSettings);
        relationships.add(footnotes);
        relationships.add(endnotes);
        relationships.add(fontTable);
        relationships.add(theme);
        if(excel){
            Node sharedStrings = new Node(RelationshipType.sharedStrings,"sharedStrings.xml","rId"+8);
            relationships.add(sharedStrings);
            for(int i = 0;i < sheets;i++){
                Node sheet = new Node(RelationshipType.sheet,"worksheets/sheet"+i+1+".xml","rId"+(8+i+1));
                relationships.add(sheet);
            }
        }
        return this;
    }
    public int getNextRid(){
        int rId = 1;
        if(relationships.size() > 0){
            String trId = relationships.stream().max(Comparator.comparing(e -> e.getRId())).get().getRId();
            rId = Integer.parseInt(trId.substring(3))+1;
        }

        return rId;
    }
    public String getNextImageName(){
        String imageName = "image1.jpeg";
        if(relationships.size() > 0){
            Node node = relationships.stream().filter(e -> e.getTarget().endsWith("jpeg")).max(Comparator.comparing(x -> x.getTarget())).orElse(null);
            if(node != null){
                String target = node.getTarget();
                int index = target.lastIndexOf(".");
                imageName = "image"+Integer.parseInt(target.substring(index-1,index))+1+".jpeg";
            }
        }

        return imageName;
    }
    public int addRelationship(RelationshipType relationshipType,String target){
        int nextRid = this.getNextRid();
        Node node = new Node(relationshipType,target,"rId"+nextRid);
        relationships.add(node);
        return nextRid;
    }

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    public class Node{
        RelationshipType relationshipType;
        String target;
        String rId;

        public CoreProperties toCoreProperties() {
            CoreProperties rela = new CoreProperties(null,"Relationship");
            rela.addAttribute("Type",relationshipType.getValue());
            rela.addAttribute("Target",this.getTarget());
            rela.addAttribute("Id",this.rId);
            return rela;
        }
    }
}
