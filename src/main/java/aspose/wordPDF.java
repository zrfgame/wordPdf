package aspose;

import com.aspose.cells.*;
import com.aspose.words.Document;
import com.aspose.words.ExportHeadersFootersMode;
import com.aspose.words.FontSettings;
import woldpdf.FontProvider;

import javax.crypto.Cipher;
import java.io.*;

/**
 * @ClassName wordPDF
 * @Description TODO
 * @Author zhurongfei
 * @Data 2020/6/22 9:45
 * Version 1.0
 **/
public class wordPDF {
    /**
     * 获取license许可凭证
     * @return
     */
    private static boolean getLicense() {
        boolean result = false;
        try {
            String licenseStr = "<License>\n"
                    + " <Data>\n"
                    + " <Products>\n"
                    + " <Product>Aspose.Total for Java</Product>\n"
                    + " <Product>Aspose.Words for Java</Product>\n"
                    + " </Products>\n"
                    + " <EditionType>Enterprise</EditionType>\n"
                    + " <SubscriptionExpiry>20991231</SubscriptionExpiry>\n"
                    + " <LicenseExpiry>20991231</LicenseExpiry>\n"
                    + " <SerialNumber>8bfe198c-7f0c-4ef8-8ff0-acc3237bf0d7</SerialNumber>\n"
                    + " </Data>\n"
                    + " <Signature>0nRuwNEddXwLfXB7pw66G71MS93gW8mNzJ7vuh3Sf4VAEOBfpxtHLCotymv1PoeukxYe31K441Ivq0Pkvx1yZZG4O1KCv3Omdbs7uqzUB4xXHlOub4VsTODzDJ5MWHqlRCB1HHcGjlyT2sVGiovLt0Grvqw5+QXBuinoBY0suX0=</Signature>\n"
                    + "</License>";
            InputStream license = new ByteArrayInputStream(licenseStr.getBytes("UTF-8"));
            License asposeLic = new License();
            asposeLic.setLicense(license);
            result = true;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }

    /**
     * word文档  转换为 PDF
     * @param inPath 源文件
     * @param outPath 目标文件
     */
    public static void doc2pdf(String inPath, String outPath) {

        //验证License，获取许可凭证
        if (!getLicense()) {
            return;
        }

        try {

            //新建一个PDF文档
            File file = new File(outPath);
            if(!file.exists()){
                file.createNewFile();
            }
            //新建一个IO输出流
            FileOutputStream os = new FileOutputStream(file);
            //获取将要被转化的word文档
            FileInputStream inputStream = new FileInputStream(new File(inPath));
            Document doc = new Document(inputStream);
//            HtmlSaveOptions hso = new HtmlSaveOptions();
//            hso.setExportRoundtripInformation(true);
//            String filePath = FontProvider.class.getClassLoader().getResource("simsunb.ttf").getFile();
//            FontSettings.getDefaultInstance().setFontsFolder(filePath, false);
            // 全面支持DOC, DOCX,OOXML, RTF HTML,OpenDocument,PDF, EPUB, XPS,SWF 相互转换
            doc.save(os, com.aspose.words.SaveFormat.PDF);
            os.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /// <summary>
    /// excel转pdf
    /// </summary>
    /// <param name="path">文件地址</param>
    /// <param name="newFilePath">转换后的文件地址</param>
    /// <returns></returns>
    public static void ExcelToPdf(String path, String newFilePath) throws Exception {
        //验证License，获取许可凭证
        if (!getLicense()) {
            return;
        }
//        FTPClient ftpClient = new FTPClient();
        Workbook wb = new Workbook(path);// 原始excel路径
        HtmlSaveOptions options = new HtmlSaveOptions();
        options.setExportDocumentProperties(false);
        options.setExportWorkbookProperties(false);
        options.setExportWorksheetProperties(false);
        options.setExportSimilarBorderStyle(true);
        options.setExportImagesAsBase64(false);
        options.setExcludeUnusedStyles(true);
        options.setExportHiddenWorksheet(false);
        options.setWidthScalable(false);
        options.setPresentationPreference(false);
        options.setHtmlCrossStringType(HtmlCrossType.CROSS_HIDE_RIGHT);
        FileOutputStream fileOS = new FileOutputStream(newFilePath);
        wb.calculateFormula();
        wb.save(fileOS, options);
        fileOS.close();
    }
    public static void main(String[] args) throws Exception {

        ExcelToPdf("C:\\Users\\98790\\Desktop\\兰考能源互联网数据接入清单.xlsx", "C:\\Users\\98790\\Desktop\\兰考能源互联网数据接入清单.html");
//        doc2pdf("C:\\Users\\98790\\Desktop\\文件管理图片上传.xlsx", "C:\\Users\\98790\\Desktop\\兰考能源互联网数据接入清单.pdf");
    }
}