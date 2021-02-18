package com.example.poidemo.word;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.officeDocument.x2006.math.CTOMath;
import org.openxmlformats.schemas.officeDocument.x2006.math.CTOMathPara;
import org.openxmlformats.schemas.officeDocument.x2006.math.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;

import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.stream.StreamResult;
import javax.xml.transform.stream.StreamSource;
import java.io.*;

/**
 * 7.创建word数学公式
 *
 * @author liugang
 * @create 2021/2/7
 */
public class CreateWordMathType {

    // 数据类型转换工厂
    private static TransformerFactory transformerFactory = TransformerFactory.newInstance();

    // 数据格式
    private static File styleSheet = new File("MML2OMML.xsl");
    private static StreamSource streamSource = new StreamSource(styleSheet);

    public static void main(String[] args) throws Exception {

        // 1.创建空文件
        XWPFDocument document = new XWPFDocument();
        // 2.创建段落
        XWPFParagraph paragraph = document.createParagraph();
        // 公式mathML
        String mathML =
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
                        + "<mrow>"
                        + "<msup><mi>a</mi><mn>2</mn></msup><mo>+</mo><msup><mi>b</mi><mn>2</mn></msup><mo>=</mo><msup><mi>c</mi><mn>2</mn></msup>"
                        + "</mrow>"
                        + "</math>";

        CTOMath ctoMath = getOMML(mathML);
        System.out.println(ctoMath);

        CTP ctp = paragraph.getCTP();
        CTOMath currentCtoMath = ctp.addNewOMath();
        currentCtoMath.set(ctoMath);

        // 文件输出
        FileOutputStream out = new FileOutputStream(new File("first-word2.docx"));
        // 写入文件
        document.write(out);
        out.close();

        System.out.println("Create First Word MathType !!!");
    }

    /**
     * 将mathML转换成omml
     *
     * @param mathML
     * @return
     */
    private static CTOMath getOMML(String mathML) throws Exception {
        // 转换器
        Transformer transformer = transformerFactory.newTransformer(streamSource);
        // mathML
        StringReader stringReader = new StringReader(mathML);
        StreamSource streamSource = new StreamSource(stringReader);
        // 转换结果放入数据流
        StringWriter stringWriter = new StringWriter();
        StreamResult result = new StreamResult(stringWriter);
        transformer.transform(streamSource, result);

        String ooml = stringWriter.toString();
        stringWriter.close();

        CTOMathPara ctoMathPara = CTOMathPara.Factory.parse(ooml);
        CTOMath ctoMath = ctoMathPara.getOMathArray(0);

        XmlCursor xmlCursor = ctoMath.newCursor();
        while (xmlCursor.hasNextToken()) {
            XmlCursor.TokenType tokenType = xmlCursor.toNextToken();
            if (tokenType.isStart()) {
                if (xmlCursor.getObject() instanceof CTR) {
                    CTR cTR = (CTR) xmlCursor.getObject();
                    cTR.addNewRPr2().addNewRFonts().setAscii("Cambria Math");
                    cTR.getRPr2().getRFonts().setHAnsi("Cambria Math");
                }
            }
        }
        return ctoMath;
    }
}
