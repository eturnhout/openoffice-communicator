package dev.evt.openoffice;

import java.util.HashMap;
import java.util.Map;

import com.sun.star.beans.PropertyValue;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.XComponent;
import com.sun.star.text.XText;
import com.sun.star.text.XTextDocument;
import com.sun.star.uno.UnoRuntime;

/**
 * <h1>TextDocument</h1>
 * <p>
 * Used to interact with text documents.
 * </p>
 *
 * @author Eelke van Turnhout <eelketurnhout3@gmail.com>
 * @version 1.0
 */
public class TextDocument extends BaseDocument
{
    public final static String NEW_DOCUMENT = "private:factory/swriter";
    public final static String PROPERTY_VALUE_FORMAT_HTML = "HTML (StarWriter)";
    public final static String PROPERTY_VALUE_FORMAT_DOC = "MS Word 97";
    public final static String PROPERTY_VALUE_FORMAT_ODF = "StarOffice XML (Writer)";
    public final static String PROPERTY_VALUE_FORMAT_PDF = "writer_pdf_Export";

    public final static String EXTENSION_DOC = ".doc";
    public final static String EXTENSION_DOCX = ".docx";
    public final static String EXTENSION_HTML = ".html";
    public final static String EXTENSION_ODF = ".odf";
    public final static String EXTENSION_PDF = ".pdf";

    public final static String EXTENSION_DEFAULT = EXTENSION_ODF;

    public final static String EXTENSIONS_AVAILABLE = EXTENSION_DOC + ", " + EXTENSION_DOCX + ", " + EXTENSION_HTML + ", " + EXTENSION_ODF + ", " + EXTENSION_PDF;

    protected final static Map<String, String> FILTER_MAPPING = initFilterMapping();

    protected XTextDocument document;

    /**
     *
     * @param connection
     *            A connection to OpenOffice.
     * @param folder
     *            Full path to an existing folder.
     * @param file
     *            (optional) the file name plus extension.
     * @param properties
     *            (optional) the document's properties. Default properties will
     *            be set when this parameter is null.
     * @throws Exception
     */
    public TextDocument(Connection connection, String folder, String file, DocumentProperties properties) throws Exception
    {
        super(connection, folder, file, properties);
        this.initialize();
    }

    /**
     * Get the document's text.
     *
     * @return the document's text in String format.
     */
    public String getText()
    {
        XText text = this.document.getText();
        return text.getString();
    }

    /**
     * Get the XTextDocument object. This can be used to interact with the
     * document.
     *
     * @return the XTextDocument object for this document.
     */
    public XTextDocument getDocument()
    {
        return this.document;
    }

    /**
     * Check if a extension is supported by the TextDocument object.
     *
     * @param extension
     *            The extension to check.
     * @return true or false.
     */
    public static boolean isValidExtension(String extension)
    {
        return FILTER_MAPPING.containsKey(extension);
    }

    /**
     * Initialize this TextDocument object. This will fill in the XTextDocument
     * object with a new or existing file if it already exists. Also sets
     * default properties when the properties parameter is set to null.
     *
     * @throws Exception
     * @throws IllegalArgumentException
     */
    private void initialize() throws Exception, IllegalArgumentException
    {
        String path = TextDocument.NEW_DOCUMENT;

        if (this.name != null) {
            if (!isValidExtension(this.extension)) {
                throw new IllegalArgumentException("Unsupported extension \"" + this.extension + "\" for object TextDocument.");
            }

            path = "file://" + this.folder + this.name + this.extension;
        } else {
            this.extension = TextDocument.EXTENSION_DEFAULT;
        }

        PropertyValue[] properties = new PropertyValue[1];
        properties[0] = new PropertyValue();
        properties[0].Name = "Hidden";
        properties[0].Value = true;

        XComponent xComponent = this.connection.getComponentLoader().loadComponentFromURL(path, "_blank", 0, properties);

        this.document = (XTextDocument) UnoRuntime.queryInterface(XTextDocument.class, xComponent);
    }

    /**
     * Maps the extensions to the correct filter type.
     *
     * @return a hash map with filters mapped to the correct extension.
     */
    protected static Map<String, String> initFilterMapping()
    {
        Map<String, String> mapping = new HashMap<String, String>();
        mapping.put(EXTENSION_DOC, PROPERTY_VALUE_FORMAT_DOC);
        mapping.put(EXTENSION_DOCX, PROPERTY_VALUE_FORMAT_DOC);
        mapping.put(EXTENSION_HTML, PROPERTY_VALUE_FORMAT_HTML);
        mapping.put(EXTENSION_ODF, PROPERTY_VALUE_FORMAT_ODF);
        mapping.put(EXTENSION_PDF, PROPERTY_VALUE_FORMAT_PDF);

        return mapping;
    }
}
