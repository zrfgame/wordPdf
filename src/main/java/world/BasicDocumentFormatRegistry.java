package world;

import com.artofsolving.jodconverter.DocumentFormat;
import com.artofsolving.jodconverter.DocumentFormatRegistry;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

class BasicDocumentFormatRegistry implements DocumentFormatRegistry {
    private List documentFormats = new ArrayList();
    public BasicDocumentFormatRegistry() {
    }
    public void addDocumentFormat(DocumentFormat documentFormat) {
        this.documentFormats.add( documentFormat );
    }
    protected List getDocumentFormats() {
        return this.documentFormats;
    }

    public DocumentFormat getFormatByFileExtension(String extension) {
        if (extension == null) {
            return null;
        } else {
            if (extension.contains("doc")) {
                extension = "doc";
            }
            if (extension.contains("ppt")) {
                extension = "ppt";
            }
            if (extension.contains("xls")) {
                extension = "xls";
            }
            String lowerExtension = extension.toLowerCase();
            Iterator it = this.documentFormats.iterator();
            DocumentFormat format;
            do {
                if (!it.hasNext()) {
                    return null;
                }
                format = (DocumentFormat) it.next();
            } while (!format.getFileExtension().equals( lowerExtension ));

            return format;
        }
    }
    public DocumentFormat getFormatByMimeType(String mimeType) {
        Iterator it = this.documentFormats.iterator();
        DocumentFormat format;
        do {
            if (!it.hasNext()) {
                return null;
            }
            format = (DocumentFormat) it.next();
        } while (!format.getMimeType().equals( mimeType ));
        return format;
    }
}
