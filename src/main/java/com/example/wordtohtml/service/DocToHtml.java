package com.example.wordtohtml.service;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.usermodel.PictureType;
import org.springframework.web.multipart.MultipartFile;
import org.w3c.dom.Document;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;

/**
 * 2003之前Word文档转HTML
 *
 * @author Peter Sun
 * @version 0.1
 * @date 2023/11/18 21:37
 */
public class DocToHtml implements IWordToHtml {

    @Override
    public void transform(final String sourceFileName) throws TransformerException {
        HWPFDocument wordDocument;
        try (FileInputStream inputStream = new FileInputStream(sourceFile)) {

            wordDocument = new HWPFDocument(inputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        // HTML文档
        Document htmlDocument = null;
        try {
            htmlDocument = DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument();
        } catch (ParserConfigurationException e) {
            throw new RuntimeException(e);
        }
        WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(htmlDocument);
        wordToHtmlConverter.setPicturesManager(new PicturesManager() {
            @Override
            public String savePicture(byte[] content, PictureType pictureType, String name, float width, float height) {
                // 图片保存到桌面
                String path = imageDir + "\\" + name;
                try (FileOutputStream out = new FileOutputStream(path)) {
                    out.write(content);
                } catch (Exception e) {
                    e.printStackTrace();
                }
                return path;
            }
        });
        // word转换为html输出到dom中
        wordToHtmlConverter.processDocument(wordDocument);

        // 获取dom转换源
        DOMSource domSource = new DOMSource(htmlDocument);
        // 获取流输出源
        StreamResult streamResult = new StreamResult(new File(targetFile));

        // java中从源到结果的转换
        TransformerFactory transformerFactory = TransformerFactory.newInstance();
        Transformer transformer = transformerFactory.newTransformer();
        // 设置属性
        transformer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
        transformer.setOutputProperty(OutputKeys.INDENT, "yes");
        transformer.setOutputProperty(OutputKeys.METHOD, "html");
        // 转换
        transformer.transform(domSource, streamResult);
    }

    @Override
    public void transform(final String sourceFileName, final String targetFileName) {

    }

    @Override
    public void transform(final MultipartFile sourceFile) {

    }

    @Override
    public void transform(final MultipartFile sourceFile, final MultipartFile targetFile) {

    }
}
