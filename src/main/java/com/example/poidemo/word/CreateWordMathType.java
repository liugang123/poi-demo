package com.example.poidemo.word;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * 7.创建word数学公式
 *
 * @author liugang
 * @create 2021/2/7
 */
public class CreateWordMathType {

    public static void main(String[] args) throws IOException {

        // 1.创建空文件
        XWPFDocument document = new XWPFDocument();

        // 2.文件输出
        FileOutputStream out = new FileOutputStream(new File("first-word.docx"));





        // 写入文件
        document.write(out);
        out.close();

        System.out.println("Create First Word MathType !!!");
    }
}
