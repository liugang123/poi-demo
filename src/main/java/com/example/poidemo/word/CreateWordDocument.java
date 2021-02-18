package com.example.poidemo.word;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * 1.创建word文档
 *
 * @author liugang
 * @create 2021/2/7
 */
public class CreateWordDocument {

    public static void main(String[] args) throws IOException {

        // 1.创建空文件
        XWPFDocument document = new XWPFDocument();

        // 2.文件输出
        FileOutputStream out = new FileOutputStream(new File("first-word.docx"));
        document.write(out);

        out.close();
        System.out.println("Create First Word Document");
    }
}
