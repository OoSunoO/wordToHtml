package com.example.wordtohtml.service;

import fr.opensagres.poi.xwpf.converter.core.ImageManager;
import fr.opensagres.poi.xwpf.converter.xhtml.XHTMLConverter;
import fr.opensagres.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;

/**
 * 2007之后Word文档转HTML
 *
 * @author Peter Sun
 * @version 0.1
 * @date 2023/11/18 21:37
 */
public class DocxToHtml implements IWordToHtml {

    @Override
    public void transform(final String sourceFileName) {
        XWPFDocument wordDocument;
        try (FileInputStream fileInputStream = new FileInputStream(sourceFile)) {

            wordDocument = new XWPFDocument(fileInputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        XHTMLOptions xhtmlOptions = XHTMLOptions.create().setImageManager(new ImageManager(new File("D:\\"), "aImage"));
        try {
            FileOutputStream fos = new FileOutputStream(targetFile);
            // 不通设备适配，有表格不行
            fos.write(("<head>" +
                    "<meta http-equiv=\"Content-Type\" content=\"text/html;charset=\"UTF-8\"/>" +
                    "<meta name=\"viewport\" content=\"initial-scale=1.0, maximum-scale=1, minimum-scale=1, user-scalable=no,uc-fitscreen=yes\" />" +
                    "<style>body {margin: 0; font-family: \"Noto Sans SC\";}" +

                    //设置图片的大小
                    "img {width: 100% !important;height: auto !important;} " +
                    "span{overflow-wrap: break-word}" +

                    //设置边距
                    "div{  width: auto !important; margin-left: 10% !important;  margin-right: 10% !important;}" +
                    "</style>" +
                    "</head>").getBytes(StandardCharsets.UTF_8));
            XHTMLConverter.getInstance().convert(wordDocument, fos, xhtmlOptions);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
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
