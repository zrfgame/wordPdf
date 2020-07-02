package world;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.ConnectException;
import java.text.SimpleDateFormat;
import java.util.Date;
import com.artofsolving.jodconverter.DocumentConverter;
import com.artofsolving.jodconverter.openoffice.connection.OpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.connection.SocketOpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.converter.OpenOfficeDocumentConverter;
import com.artofsolving.jodconverter.openoffice.converter.StreamOpenOfficeDocumentConverter;


public class Office2HtmlOrPdfUtil {

    private static Integer port=18087;
    private static String host="127.0.0.1";

    private static Office2HtmlOrPdfUtil office2HtmlOrPdfUtil;
    /** * 获取Doc2HtmlUtil实例 */
    public static synchronized Office2HtmlOrPdfUtil getDoc2HtmlUtilInstance() {
        if (office2HtmlOrPdfUtil == null) {
            office2HtmlOrPdfUtil = new Office2HtmlOrPdfUtil();
        }
        return office2HtmlOrPdfUtil;
    }
    /*** 转换文件成pdf */
    public static String file2pdf(InputStream fromFileInputStream, String toFilePath,String name,String type) throws IOException {
        Date date = new Date();
        SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss");
        String timesuffix = sdf.format(date);
        String docFileName = null;
        String htmFileName = null;
        if(".doc".equals(type)){
            docFileName = "doc_" + name + ".doc";
            htmFileName = name+ ".pdf";
        }else if(".docx".equals(type)){
            docFileName = "docx_" + timesuffix + ".docx";
            htmFileName = name+".pdf";
        }else if(".xls".equals(type)){
            docFileName = "xls_" + timesuffix + ".xls";
            htmFileName = name+".pdf";
        }else if(".ppt".equals(type)){
            docFileName = "ppt_" + timesuffix + ".ppt";
            htmFileName = name+ ".pdf";
        }else{
            return null;
        }
        File htmlOutputFile = new File(toFilePath + File.separatorChar + htmFileName);
        File docInputFile = new File(toFilePath + File.separatorChar + docFileName);
        if (htmlOutputFile.exists())
            htmlOutputFile.delete();
        htmlOutputFile.createNewFile();
        if (docInputFile.exists())
            docInputFile.delete();
        docInputFile.createNewFile();
        /*** 由fromFileInputStream构建输入文件  */
        try {
            OutputStream os = new FileOutputStream(docInputFile);
            int bytesRead = 0;
            byte[] buffer = new byte[1024 * 8];
            while ((bytesRead = fromFileInputStream.read(buffer)) != -1) {
                os.write(buffer, 0, bytesRead);
            }
            os.close();
            fromFileInputStream.close();
        } catch (IOException e) {
        }
        // 连接服务
        OpenOfficeConnection connection = new SocketOpenOfficeConnection(host,port);
        try {
            connection.connect();
        } catch (ConnectException e) {
            System.err.println("文件转换出错，请检查OpenOffice服务是否启动。");
        }
        // convert 转换
        DocumentConverter converter = new OpenOfficeDocumentConverter(connection);
        converter.convert(docInputFile, htmlOutputFile);
        connection.disconnect();
        // 转换完之后删除word文件
        docInputFile.delete();
        return htmFileName;
    }

    /**文件转换成Html*/
    public static String file2Html (InputStream fromFileInputStream, String toFilePath,String type) throws Exception{
        Date date = new Date();
        SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss");
        String timesuffix = sdf.format(date);
        String docFileName = null;
        String htmFileName = null;
        if("doc".equals(type)){
            docFileName = timesuffix.concat(".doc");
            htmFileName = timesuffix.concat(".html");
        }else if("xls".equals(type)){
            docFileName = timesuffix.concat(".xls");
            htmFileName = timesuffix.concat(".html");
        }else if("ppt".equals(type)){
            docFileName = timesuffix.concat(".ppt");
            htmFileName = timesuffix.concat(".html");
        }else if("txt".equals(type)){
            docFileName = timesuffix.concat(".txt");
            htmFileName = timesuffix.concat(".html");
        }else if("pdf".equals(type)){
            docFileName = timesuffix.concat(".pdf");
            htmFileName = timesuffix.concat(".html");
        }else{
            return null;
        }
        File htmlOutputFile = new File(toFilePath + File.separatorChar + htmFileName);
        File docInputFile = new File(toFilePath + File.separatorChar + docFileName);
        if (htmlOutputFile.exists()){
            htmlOutputFile.delete();
        }
        htmlOutputFile.createNewFile();
        docInputFile.createNewFile();
        /**
         * 由fromFileInputStream构建输入文件
         */
        int bytesRead = 0;
        byte[] buffer = new byte[1024 * 8];
        OutputStream os = new FileOutputStream(docInputFile);
        while ((bytesRead = fromFileInputStream.read(buffer)) != -1) {
            os.write(buffer, 0, bytesRead);
        }
        os.close();
        fromFileInputStream.close();
        OpenOfficeConnection connection = new SocketOpenOfficeConnection(host,port);
        connection.connect();
        // convert
        DocumentConverter converter = new StreamOpenOfficeDocumentConverter(connection);
        converter.convert(docInputFile, htmlOutputFile);
        connection.disconnect();
        // 转换完之后删除word文件
        docInputFile.delete();
        return htmFileName;
    }
    public static void main(String[] args) throws IOException {
        Office2HtmlOrPdfUtil coc2HtmlUtil = getDoc2HtmlUtilInstance ();
        File file = null;
        FileInputStream fileInputStream = null;
        String fileName = "兰考能源运行月报-5月份V1.0.docx";
        String path ="C:\\Users\\98790\\Desktop\\";
        file = new File(path+fileName);
        fileInputStream = new FileInputStream(file);
        String suffix = fileName.substring(fileName.lastIndexOf("."), fileName.length());
        String name = fileName.substring(0,fileName.lastIndexOf("."));
        String fileurl = coc2HtmlUtil.file2pdf(fileInputStream, path,name,suffix);
        System.out.println(fileurl);
    }

}
