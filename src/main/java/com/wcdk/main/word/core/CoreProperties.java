package com.wcdk.main.word.core;

import java.text.SimpleDateFormat;
import java.util.*;


public class CoreProperties implements Cloneable {
    String name = "coreProperties";
    String prefix = "cp";
    String value = "";
    List<CoreProperties> child = new ArrayList<>();
    Map<String,String> nameSpace = new HashMap<>();
    Map<String,String> attribute = new HashMap<>();


    public CoreProperties() {
    }

    public CoreProperties( String prefix,String name) {
        this.name = name;
        this.prefix = prefix;
    }
    public CoreProperties( String prefix,String name,String value) {
        this.name = name;
        this.prefix = prefix;
        this.value = value;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getPrefix() {
        return prefix;
    }

    public void setPrefix(String prefix) {
        this.prefix = prefix;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    public List<CoreProperties> getChild() {
        return child;
    }

    public Map<String, String> getNameSpace() {
        return nameSpace;
    }

    public Map<String, String> getAttribute() {
        return attribute;
    }

    public CoreProperties addNameSpace(String prefix, String uri){
        this.nameSpace.put(prefix,uri);
        return this;
    }
    public CoreProperties addAttribute(String qualifiedName, String value){
        this.attribute.put(qualifiedName,value);
        return this;
    }
    public List<CoreProperties> addChild(CoreProperties coreProperties){
        this.child.add(coreProperties);
        return this.child;
    }
    public List<CoreProperties> addChild( List<CoreProperties> coreProperties){
        this.child.addAll(coreProperties);
        return this.child;
    }
    public List<CoreProperties> addChild(CoreProperties coreProperties,CoreProperties...more){
        this.child.add(coreProperties);
        if(more != null && more.length > 0){
            this.child.addAll(Arrays.asList(more));
        }
        return this.child;
    }
    public static String getTimeString(){
        String FORMAT_T = "yyyy-MM-dd'T'HH:mm:ss'Z'";
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat(FORMAT_T);
        return simpleDateFormat.format(new Date());
    }

    @Override
    public String toString() {
        return "CoreProperties{" +
                "name='" + name + '\'' +
                ", prefix='" + prefix + '\'' +
                ", value='" + value + '\'' +
                ", child=" + child +
                ", nameSpace=" + nameSpace +
                ", attribute=" + attribute +
                '}';
    }

    @Override
    public CoreProperties clone()  {
        try {
            return (CoreProperties) super.clone();
        } catch (CloneNotSupportedException e) {
            e.printStackTrace();
        }
        return null;
    }
}
