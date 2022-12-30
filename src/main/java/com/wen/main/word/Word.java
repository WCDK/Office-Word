package com.wen.main.word;

import com.wen.main.word.core.CoreProperties;
import com.wen.main.word.core.WordItem;
import com.wen.main.word.paragraph.Paragraph;
import com.wen.main.word.table.WordTable;
import org.dom4j.*;
import org.dom4j.io.OutputFormat;
import org.dom4j.io.SAXReader;
import org.dom4j.io.XMLWriter;
import org.dom4j.tree.DefaultText;

import java.io.*;
import java.net.URI;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

public class Word {

    /***
     *
     * <w:p> <!--表示一个段落-->
     * <w:val > <!--表示一个值-->
     * <w:r> <!--表示一个样式串，指明它包括的文本的显示样式，表示一个特定的文本格式-->
     * <w:t> <!--表示真正的文本内容-->
     * <w:rPr> <!--是<w:r>标签内的标签，对Run文本属性进行修饰-->
     * <w:pPr> <!--是<w:p>标签内的标签，对Paragraph文本属性进行修饰-->
     * <w:rFronts> <!--字体-->
     * <w:hdr> <!--页眉-->
     * <w:ftr> <!--页脚-->
     * <w:drawing > <!--图片-->
     * <wp:extent> <!--绘图对象大小-->
     * <wp:effectExtent > <!--嵌入图形的效果-->
     * <wp:inline  > <!--内嵌绘图对象，dist(T,B,L，R)距离文本上下左右的距离-->
     * <w:noProof  > <!--不检查拼写和语法错误-->
     * <w:docPr> <!--表示文档属性-->
     * <w:rsidR> <!--指定唯一一个标识符，用来跟踪编辑在修订时表行标识，所有段落和段落中的内容都应该拥有相同的属性值，如果出现差异，那么表示这个段落在后面的编辑中被修改。-->
     * <w:r> <!--表示关系，段落中以相连续的中文或英文字符字符串，作为开始和结束。目的就是要把一个段落中的中英文字符区分开来。 -->
     * <w:ind> <!--w:pPr元素的子元素，跟w:pStyle并列，ind代表缩进情况：有几个属性值：①firstLine（首行缩进）②left（左缩进）③当left和firstLine同时出现时代表下面的元素有两种属性首行和下面其他行都是有属性的④hanging（悬挂）-->
     * <w:hint> <!--字体的类型，w:rFonts的子元素，属性值eastAsia表面上的意思是“东亚”，指代“中日韩CJK”类型。-->
     * <w:bCs> <!--复合字体的加粗-->
     * <w:bookmarkStart> <!--书签开始-->
     * <w:bookmarkEnd> <!--书签结束-->
     * <w:lastRenderedPageBreak > <!--页面进行分页的标记，是w:r的一个属性，表示此段字符串是一页中的最后一个字符串。-->
     * <w:smartTag > <!--智能标记-->
     * <w:attr  > <!--自定义XML属性-->
     *
     * <w:b w:val=”on”> <!--表示该格式串种的文本为粗体-->
     * <w:jc w:val="right"/> <!--表示对齐方式-->
     * <w:sz w:val="40"/> <!--表示字号大小-->
     * <w:szCs w:val="40"/> <!---->
     * <w:t xml:space="preserve"> <!--保持空格，如果没有这内容的话，文本的前后空格将会被Word忽略-->
     * <w:spacing  w:line="600" w:lineRule="auto"/> <!--设置行距，要进行运算，要用数字除以240，如此处为600/240=2.5倍行距-->
     *
     * ***/
    private CoreProperties _rels_rels;
    private CoreProperties docProps_app;
    private CoreProperties docProps_core;
    private CoreProperties word_rels_document_xml;
    private CoreProperties word_theme_theme1;
    private CoreProperties word_fontTable;
    private CoreProperties word_settings;
    private CoreProperties word_styles;
    private CoreProperties word_webSettings;
    private CoreProperties Content_Types;

    transient String BASE_PATH = "";
    transient String TEMP_PATH = "";
    transient String[] BASE_SIFFIX = {".rels", ".xml", ".jpeg", ".png"};
    transient List<String> BASE_SIFFIX_LIST;
    transient CoreProperties documentContent;


    public Word() {
        BASE_PATH = System.getProperty("java.io.tmpdir");
        BASE_SIFFIX_LIST = Arrays.asList(BASE_SIFFIX);
    }

    public void loadWord(String wordPath) throws Exception {
        Path docPath = Paths.get(wordPath);
        String s = docPath.getFileName().toString();
        String zipUrl = BASE_PATH + File.separator + s.substring(0, s.lastIndexOf(".")) + ".zip";
        Path zipPath = Paths.get(zipUrl);
        zipPath.toFile().delete();
        Files.copy(docPath, zipPath);
        unzip(zipUrl, BASE_PATH);
    }

    public Word createNewWord() {
        try {
            this.BASE_PATH = "";
            this.TEMP_PATH = "";
            this._rels_rels = fixElement("wordSource/_rels/.rels");
            this.docProps_app = fixElement("wordSource/docProps/app.xml");
            this.docProps_core = fixElement("wordSource/docProps/core.xml");
            this.word_rels_document_xml = fixElement("wordSource/word/_rels/document.xml.rels");
            this.word_theme_theme1 = fixElement("wordSource/word/theme/theme1.xml");
            this.documentContent = fixElement("wordSource/word/document.xml");
            this.word_fontTable = fixElement("wordSource/word/fontTable.xml");
            this.word_settings = fixElement("wordSource/word/settings.xml");
            this.word_styles = fixElement("wordSource/word/styles.xml");
            this.word_webSettings = fixElement("wordSource/word/webSettings.xml");
            this.Content_Types = fixElement("wordSource/[Content_Types].xml");
        }catch (Exception e){
            e.printStackTrace();
        }

        return this;
    }

    private void moveResource() {
        try {
            String sourcePath = System.getProperty("java.io.tmpdir") + File.separator;
            URL wordSource = this.getClass().getClassLoader().getResource("wordSource");
            System.out.println(wordSource.toURI());
            URI uri = wordSource.toURI();
            Path path = Paths.get(uri);
            File BaseFile = path.toFile();
            moveResource(BaseFile, sourcePath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void moveResource(File file, String sourcePath) throws Exception {
        String newPath = sourcePath + file.getName();
        File nFile = new File(newPath);
        if (file.isDirectory()) {
            if (!nFile.exists()) {
                nFile.mkdirs();
            }
            File[] files = file.listFiles();
            for (File f : files) {
                moveResource(f, newPath + File.separator);
            }
        } else {
            if (!nFile.exists()) {
                nFile.createNewFile();
            }
            Path path = Paths.get(file.toURI());
            Files.copy(path, new BufferedOutputStream(new FileOutputStream(nFile)));
        }
    }

    public void toWord(String filePath) throws Exception {
        if(!"".equals(TEMP_PATH)){
            toWordDefult(filePath);
        }else {
            toWordNew(filePath);
        }
    }

    private void toWordDefult(String filePath) throws Exception{
        ZipOutputStream zipOutputStream = new ZipOutputStream(new FileOutputStream(filePath));

        Path rootPath = Paths.get(TEMP_PATH);
        File rootFile = rootPath.toFile();
        List<String> paths = new ArrayList<>();
        parePath(rootFile, paths);
        for (String path : paths) {
            InputStream inputStream = new FileInputStream(BASE_PATH + path);
            if (path.endsWith("core.xml")) {
                SAXReader reader = new SAXReader();
                Document document = reader.read(inputStream);
                Element rootElement = document.getRootElement();
                CoreProperties properties = fixElement(rootElement);
                inputStream = propertiesToStrem(properties);
            } else if (path.endsWith("document.xml")) {
                inputStream = propertiesToStrem(this.documentContent);
            }
            path = path.substring(path.indexOf(File.separator) + 1);
            zipOutputStream.putNextEntry(new ZipEntry(path));
            if (path.contains("media")) {
                int b;
                while ((b = inputStream.read()) != -1) {
                    zipOutputStream.write(b);
                }
                inputStream.close();
            } else {
                BufferedReader bufferedReader = new BufferedReader(new InputStreamReader(inputStream, "utf-8"));
                String b;
                while ((b = bufferedReader.readLine()) != null) {
                    zipOutputStream.write(b.getBytes(StandardCharsets.UTF_8));
                }
                bufferedReader.close();
            }
        }
        zipOutputStream.flush();
        zipOutputStream.close();
    }
    private void toWordNew(String filePath) throws Exception{
        ZipOutputStream zipOutputStream = new ZipOutputStream(new FileOutputStream(filePath));
        InputStream inputStream = propertiesToStrem(this._rels_rels);
        write("_rels/.rels",inputStream,zipOutputStream);

        inputStream = propertiesToStrem(this.docProps_app);
        write("docProps/app.xml",inputStream,zipOutputStream);

        inputStream = propertiesToStrem(this.docProps_core);
        write("docProps/core.xml",inputStream,zipOutputStream);

        inputStream = propertiesToStrem(this.word_rels_document_xml);
        write("word/_rels/document.xml.rels",inputStream,zipOutputStream);

        inputStream = propertiesToStrem(this.word_theme_theme1);
        write("word/theme/theme1.xml",inputStream,zipOutputStream);

        inputStream = propertiesToStrem(this.documentContent);
        write("word/document.xml",inputStream,zipOutputStream);

        inputStream = propertiesToStrem(this.word_fontTable);
        write("word/fontTable.xml",inputStream,zipOutputStream);

        inputStream = propertiesToStrem(this.word_settings);
        write("word/settings.xml",inputStream,zipOutputStream);

        inputStream = propertiesToStrem(this.word_styles);
        write("word/styles.xml",inputStream,zipOutputStream);

        inputStream = propertiesToStrem(this.word_webSettings);
        write("word/webSettings.xml",inputStream,zipOutputStream);

        inputStream = propertiesToStrem(this.Content_Types);
        write("[Content_Types].xml",inputStream,zipOutputStream);

        zipOutputStream.flush();
        zipOutputStream.close();
    }
    private void write(String zipEntry,InputStream inputStream,ZipOutputStream outputStream) throws Exception{
        outputStream.putNextEntry(new ZipEntry(zipEntry));
        int b;
        while ((b = inputStream.read()) != -1) {
            outputStream.write(b);
        }
        inputStream.close();
    }
    private InputStream propertiesToStrem(CoreProperties properties) throws Exception {
        Document document = DocumentHelper.createDocument();

        String rootName = properties.getName();
        if (!properties.getPrefix().trim().equals("")) {
            rootName = properties.getPrefix() + ":" + rootName;
        }
        Element recordNode = document.addElement(rootName);
        Map<String, String> nameSpace = properties.getNameSpace();
        nameSpace.forEach((k, v) -> {
            recordNode.addNamespace(k, v);
        });
        Map<String, String> attribute = properties.getAttribute();
        attribute.forEach((k, v) -> {
            recordNode.addAttribute(k, v);
        });
        List<CoreProperties> child = properties.getChild();
        for (int i = 0; i < child.size(); i++) {
            CoreProperties coreProperties = child.get(i);
            stream(coreProperties, recordNode);
        }

        OutputFormat format = OutputFormat.createPrettyPrint();
        format.setEncoding("UTF-8");
        StringWriter stringWriter = new StringWriter(1024);
        XMLWriter xmlWriter = new XMLWriter(stringWriter, format);

        xmlWriter.setEscapeText(false);
        xmlWriter.write(document);
        xmlWriter.flush();
        xmlWriter.close();
        BufferedInputStream bufferedInputStream = new BufferedInputStream(new ByteArrayInputStream(stringWriter.toString().getBytes(StandardCharsets.UTF_8)));
        return bufferedInputStream;
    }

    private void stream(CoreProperties properties, Element element) {
        String name = properties.getName();
        String prefix = properties.getPrefix();
        String value = properties.getValue();
        if ("lastModifiedBy".equalsIgnoreCase(properties.getName())) {
            value = "WCDK";
        } else if ("modified".equalsIgnoreCase(properties.getName())) {
            value = properties.getTimeString();
        }

        Element element1 = element.addElement(name);
        element1.setText(value);

        Map<String, String> nameSpace = properties.getNameSpace();
        nameSpace.forEach((k, v) -> {
            element1.addNamespace(k, v);
        });
        if (!prefix.trim().equals("")) {
            name = prefix + ":" + name;
        }
        element1.setName(name);
        Map<String, String> attribute = properties.getAttribute();
        attribute.forEach((k, v) -> {
            element1.addAttribute(k, v);
        });
        List<CoreProperties> child = properties.getChild();
        for (int i = 0; i < child.size(); i++) {
            CoreProperties coreProperties = child.get(i);
            stream(coreProperties, element1);
        }
    }

    private CoreProperties fixElement(Element element) {
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

    private CoreProperties fixElement(String path) throws Exception{
        URL resource = this.getClass().getClassLoader().getResource(path);
        if(resource == null){
            return  null;
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

    private void pareElement(Element element, CoreProperties coreProperties) {
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

    private void parePath(File file, List<String> result) {
        if (file.isDirectory()) {
            File[] files = file.listFiles();
            for (File file1 : files) {
                if (file1.isDirectory()) {
                    parePath(file1, result);
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

    private void unzip(String filePath, String zipDir) {
        String name = "";
        try {
            InputStream inputStream;

            ZipEntry entry;
            ZipFile zipfile = new ZipFile(filePath);
            File file = new File(filePath);
            name = file.getName();
            name = name.substring(0, name.lastIndexOf("."));
            zipDir = zipDir + name;
            TEMP_PATH = zipDir;
            file = new File(zipDir);

            if (!file.exists()) {
                file.mkdirs();
            }
            Enumeration dir = zipfile.entries();
            while (dir.hasMoreElements()) {
                entry = (ZipEntry) dir.nextElement();
                if (entry.isDirectory()) {
                    name = entry.getName();
                    File fileObject = new File(zipDir + File.separator + name);
                    fileObject.mkdir();
                    continue;
                }

                file = new File(zipDir + File.separator + entry.getName());

                if (file.isFile() || BASE_SIFFIX_LIST.contains(entry.getName().substring(entry.getName().lastIndexOf(".")))) {
                    File parentFile = file.getParentFile();
                    if (!parentFile.exists()) {
                        parentFile.mkdirs();
                    }
                    file.delete();
                    file.createNewFile();
                }
                inputStream = zipfile.getInputStream(entry);
                if (entry.getName().endsWith("document.xml")) {
                    SAXReader reader = new SAXReader();
                    Document document = reader.read(inputStream);
                    Element rootElement = document.getRootElement();
                    CoreProperties properties = fixElement(rootElement);
                    this.documentContent = properties;
                    inputStream = propertiesToStrem(properties);
                }
                int count;
                byte[] dataByte = new byte[1024];
                try (FileOutputStream fos = new FileOutputStream(file)) {
                    while ((count = inputStream.read(dataByte, 0, 1024)) != -1) {
                        fos.write(dataByte, 0, count);
                    }
                }
                inputStream.close();
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            Path zipPath = Paths.get(filePath);
            zipPath.toFile().delete();
        }
    }

    public List<String> getAllTxt() throws Exception {
        SAXReader reader = new SAXReader();
        Document document = reader.read(new File(TEMP_PATH + File.separator + "word" + File.separator + "document.xml"));
        Element rootElement = document.getRootElement();
        return getAllTxt(rootElement);
    }

    public List<String> getParagraphAllTxt(String color) throws Exception {
        SAXReader reader = new SAXReader();
        List<String> result = new ArrayList();
        Document document = reader.read(new File(TEMP_PATH + File.separator + "word" + File.separator + "document.xml"));
        Element rootElement = document.getRootElement();
        Element body = rootElement.element("body");
        List<Element> ps = body.elements("p");
        for (Element element : ps) {
            StringBuilder stringBuilder = new StringBuilder();
            List<Element> rs = element.elements("r");

            rs.forEach(r -> {
                Element rpr = r.element("rPr");
                if (rpr != null && rpr.hasContent()) {
                    Element colorElement = rpr.element("color");
                    if (colorElement != null && colorElement.attribute("val").getStringValue().equals(color)) {
                        Element t = r.element("t");
                        String textTrim = t.getTextTrim();
                        if (!textTrim.equals("")) {
                            stringBuilder.append(textTrim);
                        }
                    }
                }

            });
            result.add(stringBuilder.toString());
        }

        return result;
    }

    private List<String> getAllTxt(Element rootElement) {
        List<String> result = new ArrayList();
        List<Node> content = rootElement.content();
        for (Node node : content) {
            if (node instanceof Namespace) {
                continue;
            } else if (node instanceof DefaultText) {
                String text = ((DefaultText) node).getText();
                if (!text.trim().equals("")) {
                    result.add(text);
                }
                continue;
            }
            Element node1 = (Element) node;
            Element t = node1.element("t");
            if (t != null) {
                result.add(t.getText());
            }
            pareText(node1, result);

        }

        return result;
    }

    private void pareText(Element element, List<String> result) {
        List<Node> content = element.content();
        for (Node node : content) {
            if (node instanceof Namespace) {
                continue;
            }
            if (node instanceof DefaultText) {
                String text = ((DefaultText) node).getText();
                if (!text.trim().equals("")) {
                    result.add(text);
                }
            } else {
                Element node1 = (Element) node;
                if (node1.hasContent()) {
                    pareText(node1, result);
                }
            }


        }
    }

    public Word append(WordItem wordItem) {
//        if (wordItem instanceof Paragraph ) {
//            return appendParagraph((Paragraph) wordItem);
//        }else if(wordItem instanceof WordTable){
//
//        }

        return appendParagraph(wordItem.toCoreProperties());
    }

    public Word appendTable(){

        return this;
    }

    public Word appendParagraph(CoreProperties context) {
        CoreProperties coreProperties = documentContent.getChild().get(0);
        List<CoreProperties> child = coreProperties.getChild();
        int index = child.size() - 1;
        CoreProperties sectPr = child.get(index);
        child.remove(index);
        coreProperties.addChild(context);
        coreProperties.addChild(sectPr);

        return this;
    }

}
