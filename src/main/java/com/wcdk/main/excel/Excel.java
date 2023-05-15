package com.wcdk.main.excel;

import com.wcdk.main.core.*;
import com.wcdk.main.excel.xl.Workbook;
import lombok.Data;
import org.dom4j.Document;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.Enumeration;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

public class Excel {
    private App app = new App();
    private Rels rels = new Rels();
    private Core core = new Core();
    private Custom custom = new Custom();
    private DocumentRels documentRels = new DocumentRels();
    private Workbook documentContentc = new Workbook();
    transient String BASE_PATH = "";
    transient String TEMP_PATH = "";
    transient List<String> BASE_SIFFIX_LIST = Arrays.asList(".rels", ".xml", ".jpeg", ".png");
    public Excel(){
        BASE_PATH = System.getProperty("java.io.tmpdir");
        documentRels.setExcel(true);

    }

    public void loadExcel(String wordPath) throws Exception {
        Path docPath = Paths.get(wordPath);
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
                    if (entry.getName().endsWith("workbook.xml")) {
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


    @Data
    public class Cell{
        Integer size;
        /** rgb **/
        String color;
        String font;
        Integer charset;
        String value;
        String space;
        String x;
        String y;
        boolean haseRpr = false;
        public CoreProperties toCoreProperties(){

            CoreProperties r = new CoreProperties(null,"r");
            CoreProperties rpr = new CoreProperties(null,"rPr");
            if(this.size != null){
                CoreProperties sz = new CoreProperties(null,"sz");
                sz.addAttribute("val",this.size.toString());
                rpr.addChild(sz);
            }
            if(this.color != null && !"".equals(this.color)){
                CoreProperties color = new CoreProperties(null,"color");
                color.addAttribute("rgb",this.color);
                rpr.addChild(color);
            }
            if(this.font != null && !"".equals(this.font)){
                CoreProperties rFont = new CoreProperties(null,"rFont");
                rFont.addAttribute("val",this.font);
                rpr.addChild(rFont);
            }
            if(this.charset != null ){
                CoreProperties charset = new CoreProperties(null,"charset");
                charset.addAttribute("val",this.color);
                rpr.addChild(charset);
            }
            CoreProperties t = new CoreProperties(null,"t",value);
            if(space != null && !"".equals(space)){
                t.addAttribute("xml:space",this.space);
            }
            if(haseRpr){
                r.addChild(rpr);
            }
            r.addChild(t);
            return r;
        }

        public void setSize(Integer size) {
            haseRpr = true;
            this.size = size;
        }

        public void setColor(String color) {
            haseRpr = true;
            this.color = color;
        }

        public void setFont(String font) {
            haseRpr = true;
            this.font = font;
        }

        public void setCharset(int charset) {
            haseRpr = true;
            this.charset = charset;
        }
    }
}
