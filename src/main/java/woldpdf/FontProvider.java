package woldpdf;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Font;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.tool.xml.XMLWorkerFontProvider;

import java.io.IOException;

/**
 * @ClassName FontProvider
 * @Description TODO
 * @Author zhurongfei
 * @Data 2020/6/16 15:05
 * Version 1.0
 **/
public class FontProvider extends XMLWorkerFontProvider {
    public Font getFont(String fontname, String encoding, boolean embedded, float size, int style, BaseColor color) {
        BaseFont bf =null;
        try {
            String filePath = FontProvider.class.getClassLoader().getResource("simsunb.ttf").getFile();
            bf = BaseFont.createFont(filePath, BaseFont.IDENTITY_H,BaseFont.NOT_EMBEDDED);//字体
        } catch (DocumentException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        Font font = this.getFont(fontname, encoding, size, style);
        font.setColor(color);
        return font;
    }
}
