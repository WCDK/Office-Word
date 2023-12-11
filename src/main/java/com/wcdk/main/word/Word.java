package com.wcdk.main.word;

import com.wcdk.main.core.*;
import com.wcdk.main.core.eunm.RelationshipType;
import com.wcdk.main.word.image.WordImage;
import com.wcdk.main.word.paragraph.Paragraph;
import com.wcdk.main.word.table.WordTable;
import org.dom4j.*;
import org.dom4j.io.OutputFormat;
import org.dom4j.io.SAXReader;
import org.dom4j.io.XMLWriter;
import org.dom4j.tree.DefaultText;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.math.BigDecimal;
import java.net.URI;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;
import java.util.stream.Collectors;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

public class Word {

    private CoreProperties word_theme_theme1;
    private CoreProperties word_fontTable;
    private CoreProperties word_endnotes;
    private CoreProperties word_footnotes;
    private CoreProperties word_settings;
    private CoreProperties word_styles;
    private CoreProperties word_webSettings;
    private CoreProperties Content_Types;
    private List<WordImage> wordImages = new ArrayList<>();
    private App app = new App();
    private Rels rels = new Rels();
    private Core core = new Core();
    private Custom custom = new Custom();
    private DocumentRels documentRels = new DocumentRels();
    private Numbering numbering = new Numbering();

    transient String BASE_PATH = "";
    transient String TEMP_PATH = "";
    transient String[] BASE_SIFFIX = {".rels", ".xml", ".jpeg", ".png"};
    transient List<String> BASE_SIFFIX_LIST;
    transient CoreProperties documentContentc;
    transient DocumentContent documentContent;
    boolean sort = false;
    private String sortIlvl = "0";
    String numId = "1";


    public Word() {
        BASE_PATH = System.getProperty("java.io.tmpdir");
        BASE_SIFFIX_LIST = Arrays.asList(BASE_SIFFIX);
    }

    public void loadWord(String wordPath) throws Exception {
        Path docPath = Paths.get(wordPath);
//        String s = docPath.getFileName().toString();
        long l = System.currentTimeMillis();
        String zipUrl = BASE_PATH + File.separator + l + ".zip";
        Path zipPath = Paths.get(zipUrl);
        zipPath.toFile().delete();
        Files.copy(docPath, zipPath);
        unzip(zipUrl, BASE_PATH);
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
                if (!entry.getName().endsWith("jpeg")) {
                    SAXReader reader = new SAXReader();
                    Document document = reader.read(inputStream);
                    Element rootElement = document.getRootElement();
                    if (entry.getName().endsWith("document.xml")) {
                        CoreProperties properties = OfficeUtil.fixElement(rootElement);
                        this.documentContentc = properties;
                        inputStream = zipfile.getInputStream(entry);
                        int count;
                        byte[] dataByte = new byte[1024];
                        try (FileOutputStream fos = new FileOutputStream(file)) {
                            while ((count = inputStream.read(dataByte, 0, 1024)) != -1) {
                                fos.write(dataByte, 0, count);
                            }
                        }
                        inputStream.close();
                    } else if (entry.getName().endsWith("app.xml")) {
                        this.app = new App(rootElement);
                    } else if (entry.getName().equals("_rels/.rels")) {
                        this.rels = new Rels(rootElement);
                    } else if (entry.getName().endsWith("core.xml")) {
                        this.core = new Core(rootElement);
                    } else if (entry.getName().endsWith("custom.xml")) {
                        this.custom = new Custom(rootElement);
                    } else if (entry.getName().endsWith("document.xml.rels")) {
                        this.documentRels = new DocumentRels(rootElement);
                    }else if (entry.getName().endsWith("numbering.xml")) {
                        this.numbering = new Numbering(rootElement);
                        this.sort = true;
                    } else {
                        inputStream = zipfile.getInputStream(entry);
                        int count;
                        byte[] dataByte = new byte[1024];
                        try (FileOutputStream fos = new FileOutputStream(file)) {
                            while ((count = inputStream.read(dataByte, 0, 1024)) != -1) {
                                fos.write(dataByte, 0, count);
                            }
                        }
                        inputStream.close();
                    }

                } else {
                    int count;
                    byte[] dataByte = new byte[1024];
                    try (FileOutputStream fos = new FileOutputStream(file)) {
                        while ((count = inputStream.read(dataByte, 0, 1024)) != -1) {
                            fos.write(dataByte, 0, count);
                        }
                    }
                    inputStream.close();
                }

            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            Path zipPath = Paths.get(filePath);
            zipPath.toFile().delete();
        }
    }
    public Word createNewWord(Boolean sort,String sortNumId) {
        this.sort = sort;
        this.numId = sortNumId;
        this.numbering.setUsed(sort);
        return createNewWord();
    }
    public Word createNewWord() {
        try {
            this.BASE_PATH = "";
            this.TEMP_PATH = "";
            this.rels = this.rels.create();
            this.custom = this.custom.create();
            this.documentRels = this.documentRels.create();
            this.documentContent = new DocumentContent();
            this.word_theme_theme1 = OfficeUtil.fixElement("wordSource/theme1.xml");
            this.word_fontTable = OfficeUtil.fixElement("wordSource/fontTable.xml");
            this.word_endnotes = OfficeUtil.fixElement("wordSource/endnotes.xml");
            this.word_footnotes = OfficeUtil.fixElement("wordSource/footnotes.xml");
            this.word_settings = OfficeUtil.fixElement("wordSource/settings.xml");
            this.word_styles = OfficeUtil.fixElement("wordSource/styles.xml");
            this.word_webSettings = OfficeUtil.fixElement("wordSource/webSettings.xml");
            this.Content_Types = OfficeUtil.fixElement("wordSource/[Content_Types].xml");
            if(sort){
                this.documentRels.addRelationship(RelationshipType.numbering,"numbering.xml");
                CoreProperties coreProperties = new CoreProperties(null,"Override");
                coreProperties.addAttribute("PartName","/word/numbering.xml");
                coreProperties.addAttribute("ContentType","application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml");
                this.Content_Types.addChild(coreProperties);
            }


        } catch (Exception e) {
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
        if (!"".equals(TEMP_PATH)) {
            toWordDefult(filePath);
        } else {
            toWordNew(filePath);
        }
    }

    private void toWordDefult(String filePath) throws Exception {
        ZipOutputStream zipOutputStream = new ZipOutputStream(new FileOutputStream(filePath));

        Path rootPath = Paths.get(TEMP_PATH);
        File rootFile = rootPath.toFile();
        List<String> paths = new ArrayList<>();
        OfficeUtil.parePath(rootFile, paths,BASE_PATH);
        for (String path : paths) {
            InputStream inputStream = new FileInputStream(BASE_PATH + path);
            if (path.endsWith("core.xml")) {
                inputStream = propertiesToStrem(this.core.toCoreProperties());
            } else if (path.endsWith("document.xml")) {
                inputStream = propertiesToStrem(this.documentContentc);
            } else if (path.endsWith("app.xml")) {
                inputStream = propertiesToStrem(this.app.toCoreProperties());
            } else if (path.endsWith("document.xml.rels")) {
                inputStream = propertiesToStrem(this.documentRels.toCoreProperties());
            } else if (path.endsWith(".rels")) {
                inputStream = propertiesToStrem(this.rels.toCoreProperties());
            } else if (path.endsWith("custom.xml")) {
                inputStream = propertiesToStrem(this.custom.toCoreProperties());
            }
            if(numbering.isUsed()){
                inputStream = propertiesToStrem(this.numbering.toCoreProperties());
                write("word/numbering.xml", inputStream, zipOutputStream);
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

    private void toWordNew(String filePath) throws Exception {
        ZipOutputStream zipOutputStream = new ZipOutputStream(new FileOutputStream(filePath));

        InputStream inputStream = propertiesToStrem(this.rels.toCoreProperties());
        write("_rels/.rels", inputStream, zipOutputStream);

        inputStream = propertiesToStrem(this.app.toCoreProperties());
        write("docProps/app.xml", inputStream, zipOutputStream);

        inputStream = propertiesToStrem(this.core.toCoreProperties());
        write("docProps/core.xml", inputStream, zipOutputStream);

        inputStream = propertiesToStrem(this.custom.toCoreProperties());
        write("docProps/custom.xml", inputStream, zipOutputStream);

        inputStream = propertiesToStrem(this.documentRels.toCoreProperties());
        write("word/_rels/document.xml.rels", inputStream, zipOutputStream);

        inputStream = propertiesToStrem(this.word_theme_theme1);
        write("word/theme/theme1.xml", inputStream, zipOutputStream);

        inputStream = propertiesToStrem(this.documentContent.toCoreProperties());
        write("word/document.xml", inputStream, zipOutputStream);

        inputStream = propertiesToStrem(this.word_fontTable);
        write("word/fontTable.xml", inputStream, zipOutputStream);

        inputStream = propertiesToStrem(this.word_endnotes);
        write("word/endnotes.xml", inputStream, zipOutputStream);

        inputStream = propertiesToStrem(this.word_footnotes);
        write("word/footnotes.xml", inputStream, zipOutputStream);

        inputStream = propertiesToStrem(this.word_settings);
        write("word/settings.xml", inputStream, zipOutputStream);

        inputStream = propertiesToStrem(this.word_styles);
        write("word/styles.xml", inputStream, zipOutputStream);

        inputStream = propertiesToStrem(this.word_webSettings);
        write("word/webSettings.xml", inputStream, zipOutputStream);

        inputStream = propertiesToStrem(this.Content_Types);
        write("[Content_Types].xml", inputStream, zipOutputStream);
        if(numbering.isUsed()){
            inputStream = propertiesToStrem(this.numbering.toCoreProperties());
            write("word/numbering.xml", inputStream, zipOutputStream);
        }

        if (wordImages.size() > 0) {
            wordImages.forEach(e -> {
                try {
                    InputStream iin = new BufferedInputStream(new FileInputStream(e.getPicSrc()));
                    write("word/" + e.getTarget(), iin, zipOutputStream);
                } catch (Exception ex) {
                    ex.printStackTrace();
                }
            });
        }

        zipOutputStream.flush();
        zipOutputStream.close();
    }

    private void write(String zipEntry, InputStream inputStream, ZipOutputStream outputStream) throws Exception {
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
        if (properties.getPrefix() != null && !properties.getPrefix().trim().equals("")) {
            rootName = properties.getPrefix() + ":" + rootName;
        }
        Map<String, String> nameSpace = properties.getNameSpace();
        String xmlns = nameSpace.get("");
        Element cuNode = null;
        if (xmlns != null) {
            cuNode = document.addElement(rootName, xmlns);
            nameSpace.remove("");
        } else {
            cuNode = document.addElement(rootName);
        }
        Element recordNode = cuNode;
        if (nameSpace != null && nameSpace.size() > 0) {
            nameSpace.forEach((k, v) -> {
                recordNode.addNamespace(k, v);
            });
        }
        Map<String, String> attribute = properties.getAttribute();
        attribute.forEach((k, v) -> {
            recordNode.addAttribute(k, v);
        });
        List<CoreProperties> child = properties.getChild();
        for (int i = 0; i < child.size(); i++) {
            CoreProperties coreProperties = child.get(i);
            OfficeUtil.stream(coreProperties, recordNode, xmlns);
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

    private Document propertiesToDocument(CoreProperties properties) throws Exception {
        Document document = DocumentHelper.createDocument();

        String rootName = properties.getName();
        if (properties.getPrefix() != null && !properties.getPrefix().trim().equals("")) {
            rootName = properties.getPrefix() + ":" + rootName;
        }
        Map<String, String> nameSpace = properties.getNameSpace();
        String xmlns = nameSpace.get("");
        Element cuNode = null;
        if (xmlns != null) {
            cuNode = document.addElement(rootName, xmlns);
            nameSpace.remove("");
        } else {
            cuNode = document.addElement(rootName);
        }
        Element recordNode = cuNode;
        if (nameSpace != null && nameSpace.size() > 0) {
            nameSpace.forEach((k, v) -> {
                recordNode.addNamespace(k, v);
            });
        }
        Map<String, String> attribute = properties.getAttribute();
        attribute.forEach((k, v) -> {
            recordNode.addAttribute(k, v);
        });
        List<CoreProperties> child = properties.getChild();
        for (int i = 0; i < child.size(); i++) {
            CoreProperties coreProperties = child.get(i);
            OfficeUtil.stream(coreProperties, recordNode, xmlns);
        }
        return document;
    }



    public List<String> getAllTxt() throws Exception {
        CoreProperties coreProperties = this.documentContent ==null?this.documentContentc:documentContent.toCoreProperties();
        Document document = propertiesToDocument(coreProperties);
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

    public Word append(OfficeItem wordItem,OfficeItem...more) {
        if (wordItem instanceof WordImage image) {
            addImage(image);
            Paragraph paragraph = new Paragraph();
            if(image.getAlgin() != null){
                paragraph.setAlgin(image.getAlgin());
            }
            paragraph.setWordItem(wordItem);
            appendParagraph(paragraph);
        } else if (wordItem instanceof Paragraph p) {
            if(p.getText() != null){
                int length = p.getText().trim().length();
                app.setWords(app.getWords() + length);
                app.setCharacters(p.getText().length());
                app.setCharacterswithspaces(p.getText().length());
            }
            appendParagraph(wordItem);
        }
        if(more != null && more.length > 0){
            Arrays.asList(more).forEach(e->{
                if (e instanceof WordImage image) {
                    addImage(image);
                    Paragraph paragraph = new Paragraph();
                    if(image.getAlgin() != null){
                        paragraph.setAlgin(image.getAlgin());
                    }
                    paragraph.setWordItem(e);
                    appendParagraph(paragraph);
                } else if (e instanceof Paragraph p) {
                    if(p.getText() != null){
                        int length = p.getText().trim().length();
                        app.setWords(app.getWords() + length);
                        app.setCharacters(p.getText().length());
                        app.setCharacterswithspaces(p.getText().length());
                    }
                    appendParagraph(e);
                }
            });
        }
        return this;
    }

    private void addImage(WordImage image) {

        String newTarget = "media/" + this.documentRels.getNextImageName();
        int id = this.documentRels.addRelationship(RelationshipType.image, newTarget);

        image.setTarget(newTarget);
//        image.setId(String.valueOf(id));
        image.setId(String.valueOf(wordImages.size()));
        image.setName("图片 " + id);
        image.setEmbed("rId" + id);

        try {
            File file = new File(image.getPicSrc());
            BufferedImage bufferedImage = ImageIO.read(new FileInputStream(file));
            int width = bufferedImage.getWidth();
            int height = bufferedImage.getHeight();
            double x = width * 9525;
            double y = height * 9525;
            int cx = BigDecimal.valueOf(x).intValue();
            int cy = BigDecimal.valueOf(y).intValue();
            cx = cx >= 5274310 ? 5274310 : cx;
            cy = cy >= 3110865 ? 3110865 : cy;
            String descr = file.getName();
            descr = descr.substring(0, descr.lastIndexOf("."));
            image.setDescr(descr);
            image.setExtent_cx(cx);
            image.setExtent_cy(cy);
            image.setExtent_r(2540);
            bufferedImage.flush();
            file = null;
        } catch (Exception e) {
            e.printStackTrace();
        }


        this.wordImages.add(image);

    }

    public Word getTable() {
        List<OfficeItem> wordItems = this.documentContent.getWordItems();
        List<CoreProperties> child = new ArrayList<>();
        wordItems.forEach(wordItem->{
            if(wordItem instanceof WordTable){
                child.add(wordItem.toCoreProperties());
            }
        });
//        List<CoreProperties> child = this.documentContent.getChild();
        for (CoreProperties coreProperties : child) {
            String name = coreProperties.getName();
            if (name.equals("tbl")) {
                List<CoreProperties> tableChild = coreProperties.getChild();
                int rowNum = tableChild.size() - 2;
                int cellNum = tableChild.get(1).getChild().size();

                WordTable wordTable = new WordTable(rowNum, cellNum);
            }
        }
        return this;
    }


    public Word appendParagraph(OfficeItem wordItem) {
        if(sort && wordItem instanceof Paragraph paragraph){
            paragraph.setSort(true);
            paragraph.setNumId(numId);
            int i = Integer.parseInt(this.sortIlvl)+1;
            paragraph.setSortIlvl(i+"");
            this.sortIlvl = i+"";
            this.numbering.buildLvl(paragraph.getStyle());
        }
         this.documentContent.getWordItems().add(wordItem);
        return this;
    }
    /**
     * <p>统计word字数 不统计空格</p>
     *
     * @Description:
     * @Author: wench
     **/

    public int countWords() {
        try {
            List<String> allTxt = getAllTxt();
            String collect = allTxt.stream().collect(Collectors.joining());
            int length = collect.length();
            this.app.setCharacters(length);
            this.app.setCharacterswithspaces(length);
            int length1 = collect.replaceAll(" ", "").length();
            this.app.setWords(length1);
        } catch (Exception e) {
            e.printStackTrace();
        }

        return this.app.getWords();
    }

    public int countPages() {
        return this.app.getPages();
    }

    public int countLines() {
        return this.app.getLines();
    }

    public int countParagraphs() {
        return this.app.getParagraphs();
    }

}
