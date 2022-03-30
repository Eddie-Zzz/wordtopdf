package com.company.html.convert;

import com.company.common.enums.FontEnum;
import com.lowagie.text.pdf.BaseFont;
import org.xhtmlrenderer.pdf.ITextFontResolver;
import org.xhtmlrenderer.pdf.ITextRenderer;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public class HtmlConvertToPdf {

//    public static void main(String[] args) throws Exception {
//        convertHtmlToPdf("aaa.html", "aaa.pdf", "D:/test");
//    }

    /**
     * html转换成pdf
     *
     * @param inputFileName 待转换的文档名称
     * @param outputFileName 输出的文档名称
     * @param catalogue 操作目录
     * @throws Exception
     */
    public static void convertHtmlToPdf(String inputFileName, String outputFileName, String catalogue) {
        OutputStream os = null;
        try {
            String inputFile = catalogue + File.separator + inputFileName;
            String outputFile = catalogue + File.separator + outputFileName;

            os = new FileOutputStream(outputFile);
            ITextRenderer renderer = new ITextRenderer();
            String url = new File(inputFile).toURI().toURL().toString();
            renderer.setDocument(url);
            // 解决中文支持问题
            ITextFontResolver fontResolver = renderer.getFontResolver();
            // windows 环境 添加字体
            fontResolver.addFont(FontEnum.SIM_SUN.getPath(), BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
            //解决图片的相对路径问题
            renderer.getSharedContext().setBaseURL("file:/" + catalogue +"/");
            renderer.layout();
            renderer.createPDF(os);
            os.flush();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (os != null) {
                try {
                    os.close();
                }catch (IOException e) {}
            }
        }
    }

}
