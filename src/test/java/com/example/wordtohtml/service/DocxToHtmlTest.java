package com.example.wordtohtml.service;

import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import javax.xml.transform.TransformerException;
import java.io.File;

import static com.example.wordtohtml.service.IWordToHtml.imageDir;

/**
 * @author Peter Sun
 * @version 0.1
 * @date 2023/11/18 23:13
 */
class DocxToHtmlTest {

    @BeforeEach
    void setUp() {
        File dir = new File(imageDir);
        if (!dir.exists()) {
            dir.mkdir();
        }
    }

    @Test
    void transform() throws TransformerException {
        IWordToHtml tool = new DocxToHtml();
        tool.transform(imageDir);
    }
}
