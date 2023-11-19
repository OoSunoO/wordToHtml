package com.example.wordtohtml;

import com.example.wordtohtml.service.DocToHtml;
import com.example.wordtohtml.service.DocxToHtml;
import com.example.wordtohtml.service.IWordToHtml;

import javax.xml.transform.TransformerException;
import java.io.File;

import static com.example.wordtohtml.service.IWordToHtml.imageDir;

/**
 * 调用方法:
 *
 * @author Peter Sun
 * @version 0.1
 * @date 2023/11/19 20:31
 */
public class WordToHtml {
    public static void main(String[] args) throws TransformerException {
        File dir = new File(imageDir);
        if (!dir.exists()) {
            dir.mkdir();
        }
        if (".docx".equals(args[0])) {
            IWordToHtml tool = new DocxToHtml();
            tool.transform(imageDir);
        } else {
            IWordToHtml tool = new DocToHtml();
            tool.transform(imageDir);
        }
    }
}
