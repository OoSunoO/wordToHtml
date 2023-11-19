package com.example.wordtohtml.service;

import org.springframework.web.multipart.MultipartFile;

import javax.xml.transform.TransformerException;

/**
 * Word转HTML接口
 *
 * @author Peter Sun
 * @version 0.1
 * @date 2023/11/18 21:43
 */
public interface IWordToHtml {
    String sourceFile = "D:\\上海市虹口区2022年政府信息公开工作年度报告.doc";
    String targetFile = "D:\\a.html";
    String imageDir = "D:\\aImg";

    /**
     * @param sourceFileName 需要转换的文件名（包括后缀）
     */
    void transform(final String sourceFileName) throws TransformerException;

    void transform(final String sourceFileName, final String targetFileName);

    void transform(final MultipartFile sourceFile);

    void transform(final MultipartFile sourceFile, final MultipartFile targetFile);
}
