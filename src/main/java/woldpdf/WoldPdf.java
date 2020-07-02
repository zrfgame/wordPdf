package woldpdf;

import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Font;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.tool.xml.XMLWorkerHelper;
import fr.opensagres.xdocreport.utils.StringUtils;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.usermodel.PictureType;
import org.apache.poi.util.StringUtil;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.w3c.dom.Document;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.*;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;
import java.nio.charset.Charset;

import static javax.xml.transform.OutputKeys.ENCODING;

/**
 * @ClassName WoldPdf
 * @Description TODO
 * @Author zhurongfei
 * @Data 2020/6/16 10:52
 * Version 1.0
 **/
public class WoldPdf {
    public static void main(String[] args) throws ParserConfigurationException, TransformerException, IOException {
        WoldPdf woldPdf = new WoldPdf();
        woldPdf.WorldHtml();
    }

    /***
     * word转html
     * @throws IOException
     * @throws ParserConfigurationException
     * @throws TransformerException
     */
    public void WorldHtml() throws IOException, ParserConfigurationException, TransformerException {
        final String filepath = "C:/Users/98790/Desktop/";
        String doFile = filepath + "兰考源码分析报告_version2.0.docx";
        final String picturesPath = filepath + "image/";
        File picturesDir = new File(picturesPath);
        String content = null;
        HWPFDocument wordDocument = new HWPFDocument(new FileInputStream(doFile));
        WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());
        wordToHtmlConverter.setPicturesManager(new PicturesManager() {
            public String savePicture(byte[] content, PictureType pictureType, String suggestedName, float widthInches, float heightInches) {
                File file = new File(picturesPath + suggestedName);
                FileOutputStream fos = null;
                try {
                    fos = new FileOutputStream(file);
                    fos.write(content);
                    fos.close();
                } catch (Exception e) {
                    e.printStackTrace();
                }
                return picturesPath + suggestedName;
            }
        });
        wordToHtmlConverter.processDocument(wordDocument);
        Document htmlDocument = wordToHtmlConverter.getDocument();
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        DOMSource domSource = new DOMSource(htmlDocument);
        StreamResult streamResult = new StreamResult(out);
        TransformerFactory tf = TransformerFactory.newInstance();
        Transformer serializer = tf.newTransformer();
        serializer.setOutputProperty(ENCODING, ENCODING);
        serializer.setOutputProperty(OutputKeys.INDENT, "yes");
        serializer.setOutputProperty(OutputKeys.METHOD, "html");
        serializer.transform(domSource, streamResult);
        out.close();
        content = out.toString();
        System.out.println(content);
        jsoupHtml(content);
    }

    /**
     * 保存到磁盘中
     *
     * @param content
     * @throws IOException
     */
    public void jsoupHtml(String content) throws IOException {
        FileOutputStream fos = null;
        BufferedWriter bw = null;
        org.jsoup.nodes.Document doc = Jsoup.parse(content);
//        content = doc.html();
        String style = doc.body().attr("style");
        if(StringUtils.isNotEmpty(style)&&style.indexOf("width")!=-1){
            doc.body().attr("style","");
        }
        Elements divs = doc.select("div");
        for(int i=0;i<divs.size();i++){
            Element div = divs.get(i);
            style = div.attr("style");
            if(StringUtils.isNotEmpty(style) && style.indexOf("width")!=-1){
                div.attr("style","");
            }
        }
        try {
            File file = new File("C:\\Users\\98790\\Desktop\\a.html");
            if (!file.exists()) {
                file.createNewFile();
            }
            fos = new FileOutputStream(file);
            bw = new BufferedWriter(new OutputStreamWriter(fos, "UTF-8"));
            bw.write(content);
            htmlPdf(file);
        } catch (Exception e) {
            e.printStackTrace();
            ;
        } finally {
            if (bw != null)
                bw.close();
            if (fos != null)
                fos.close();
        }

    }

    /**
     * html转成pdf
     *
     * @throws IOException
     * @throws DocumentException
     */
    public void htmlPdf(File htmlFile) throws IOException, DocumentException {
        String outputPdfPath = "C:\\Users\\98790\\Desktop\\a.pdf";
        File file = new File(outputPdfPath);
        if (!file.exists()) {
            file.createNewFile();
        }
        com.itextpdf.text.Document document = new com.itextpdf.text.Document();
        PdfWriter writer = PdfWriter.getInstance(document, new FileOutputStream(outputPdfPath));
        document.open();
        String filePath = FontProvider.class.getClassLoader().getResource("simsunb.ttf").getFile();
        BaseFont bf = BaseFont.createFont("STSong-Light", BaseFont.IDENTITY_H,BaseFont.NOT_EMBEDDED);//字体
        Font font = new Font(bf);
        document.add(new Paragraph("解决中文问题",font));
        XMLWorkerHelper.getInstance().parseXHtml(writer, document, new FileInputStream(htmlFile), Charset.forName("UTF-8"));
        document.close();
    }
}