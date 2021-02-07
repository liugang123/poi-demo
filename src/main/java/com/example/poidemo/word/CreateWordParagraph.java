package com.example.poidemo.word;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * 2.创建word段落
 *
 * @author liugang
 * @create 2021/2/7
 */
public class CreateWordParagraph {

    public static void main(String[] args) throws IOException {

        // 1.创建word文档
        XWPFDocument document = new XWPFDocument();

        // 2.文件输出
        FileOutputStream out = new FileOutputStream(new File("first-word.docx"));

        // 3.创建标题段落
        XWPFParagraph titleParagraph = document.createParagraph();
        //设置段落居中
        titleParagraph.setAlignment(ParagraphAlignment.CENTER);

        // 3.1 居中样式
        XWPFRun titleRun = titleParagraph.createRun();
        titleRun.setText("文档标题");
        titleRun.setColor("000000");
        titleRun.setFontSize(20);

        // 3.创建段落
        XWPFParagraph paragraph = document.createParagraph();
        // 3.1 创建行
        XWPFRun run = paragraph.createRun();
        run.setText("举头望明月");

        // 4.写入文件
        document.write(out);
        out.close();

        System.out.println("Create First Word Paragraph !!!");
    }
}
