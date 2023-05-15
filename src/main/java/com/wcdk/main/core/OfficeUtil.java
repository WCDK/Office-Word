package com.wcdk.main.core;

import org.dom4j.*;
import org.dom4j.io.SAXReader;
import org.dom4j.tree.DefaultText;

import java.io.File;
import java.io.InputStream;
import java.net.URL;
import java.util.List;
import java.util.Map;

public class OfficeUtil {

    public static void stream(CoreProperties properties, Element element, String xmlns) {
        String name = properties.getName();
        String prefix = properties.getPrefix();
        String value = properties.getValue();
        Map<String, String> nameSpace = properties.getNameSpace();


        Element elementt = null;
        if (xmlns != null) {
            elementt = element.addElement(name, xmlns);
        } else {
            elementt = element.addElement(name);
        }
        if (prefix != null && !prefix.trim().equals("")) {
            name = prefix + ":" + name;
            elementt.setName(name);
        }
        Element element1 = elementt;
        if (nameSpace != null && nameSpace.size() > 0) {
            nameSpace.forEach((k, v) -> {
                element1.addNamespace(k, v);
            });
        }


        Map<String, String> attribute = properties.getAttribute();
        if (attribute != null && attribute.size() > 0) {
            attribute.forEach((k, v) -> {
                element1.addAttribute(k, v);
            });
        }
        if (value != null) {
            element1.setText(value);
        }
        List<CoreProperties> child = properties.getChild();
        for (int i = 0; i < child.size(); i++) {
            CoreProperties coreProperties = child.get(i);
            stream(coreProperties, element1, xmlns);
        }
    }

    public static CoreProperties fixElement(Element element) {
        CoreProperties coreProperties = new CoreProperties();
        coreProperties.setName(element.getName());
        coreProperties.setPrefix(element.getNamespacePrefix());
        List<Namespace> namespaces = element.declaredNamespaces();
        for (Namespace namespace : namespaces) {
            coreProperties.addNameSpace(namespace.getPrefix(), namespace.getURI());
        }
        List<Attribute> attributes = element.attributes();
        for (Attribute attribute : attributes) {
            coreProperties.addAttribute(attribute.getQualifiedName(), attribute.getValue());
        }
        List<Node> content = element.content();
        for (Node node : content) {
            if (node instanceof Namespace) {
                continue;
            } else if (node instanceof DefaultText) {
                if (node.getName() == null) {
                    continue;
                }
                DefaultText defaultText = (DefaultText) node;
                CoreProperties chi = new CoreProperties();
                chi.setName(defaultText.getName());
                chi.setValue(defaultText.getText());
                coreProperties.addChild(chi);
                continue;
            }
            Element nodeDocument = (Element) node;
            pareElement(nodeDocument, coreProperties);
        }
        return coreProperties;
    }

    public static CoreProperties fixElement(String path) throws Exception {
        URL resource = OfficeUtil.class.getClassLoader().getResource(path);
        if (resource == null) {
            return null;
        }
        InputStream inputStream = resource.openStream();
        SAXReader reader = new SAXReader();
        Document document = reader.read(inputStream);
        Element element = document.getRootElement();
        CoreProperties coreProperties = new CoreProperties();
        coreProperties.setName(element.getName());
        coreProperties.setPrefix(element.getNamespacePrefix());
        List<Namespace> namespaces = element.declaredNamespaces();
        for (Namespace namespace : namespaces) {
            coreProperties.addNameSpace(namespace.getPrefix(), namespace.getURI());
        }
        List<Attribute> attributes = element.attributes();
        for (Attribute attribute : attributes) {
            coreProperties.addAttribute(attribute.getQualifiedName(), attribute.getValue());
        }
        List<Node> content = element.content();
        for (Node node : content) {
            if (node instanceof Namespace) {
                continue;
            } else if (node instanceof DefaultText) {
                if (node.getName() == null) {
                    continue;
                }
                DefaultText defaultText = (DefaultText) node;
                CoreProperties chi = new CoreProperties();
                chi.setName(defaultText.getName());
                chi.setValue(defaultText.getText());
                coreProperties.addChild(chi);
                continue;
            }
            Element nodeDocument = (Element) node;
            pareElement(nodeDocument, coreProperties);
        }
        return coreProperties;
    }

    public static void pareElement(Element element, CoreProperties coreProperties) {
        CoreProperties child = new CoreProperties();
        child.setPrefix(element.getNamespacePrefix());
        child.setName(element.getName());
        List<Attribute> attributes = element.attributes();
        for (Attribute attribute : attributes) {
            child.addAttribute(attribute.getQualifiedName(), attribute.getValue());
        }
        List<Namespace> namespaces = element.declaredNamespaces();
        for (Namespace namespace : namespaces) {
            child.addNameSpace(namespace.getPrefix(), namespace.getURI());
        }
        coreProperties.addChild(child);

        List<Node> content = element.content();
        for (Node node : content) {
            if (node instanceof Namespace) {
                continue;
            }
            if (node instanceof DefaultText) {
                DefaultText defaultText = (DefaultText) node;
                child.setValue(defaultText.getText());
            } else if (node instanceof Element) {
                Element childElement = (Element) node;
                pareElement(childElement, child);
            }
        }
    }

    public static void parePath(File file, List<String> result,String BASE_PATH) {
        if (file.isDirectory()) {
            File[] files = file.listFiles();
            for (File file1 : files) {
                if (file1.isDirectory()) {
                    parePath(file1, result,BASE_PATH);
                } else {
                    String path = file1.getPath();
                    String substring = path.substring(BASE_PATH.length());
                    result.add(substring);
                }
            }
        } else {
            String path = file.getPath();
            String substring = path.substring(BASE_PATH.length());
            result.add(substring);
        }
    }
}
