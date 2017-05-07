package dev.evt.openoffice;

import com.sun.star.beans.PropertyValue;
import com.sun.star.io.IOException;
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

    public final static String EXTENSION_DOC = ".doc";
    public final static String EXTENSION_DOCX = ".docx";
    public final static String EXTENSION_HTML = ".html";

    public final static String EXTENSIONS_AVAILABLE = EXTENSION_DOC + ", " + EXTENSION_DOCX + ", " + EXTENSION_HTML;

    protected XTextDocument document;

    public TextDocument(Connection connection, String folder, String file, DocumentProperties properties)
            throws IOException, IllegalArgumentException
    {
        super(connection, folder, file, properties);
        this.initialize();
    }

    /**
     * Get the documents text
     * 
     * @return the documents text in String format
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
     * @return the XTextDocument object for this document
     */
    public XTextDocument getDocument()
    {
        return this.document;
    }

    /**
     * Initialize this object. This will fill in the XTextDocument object with a
     * new or existing file if it already exists
     *
     * @throws IOException
     * @throws IllegalArgumentException
     */
    private void initialize() throws IOException, IllegalArgumentException
    {
        String path = TextDocument.NEW_DOCUMENT;

        if (this.name != null) {
            path = "file://" + this.folder + this.name + this.extension;
        }

        PropertyValue[] lProperties = new PropertyValue[1];
        lProperties[0] = new com.sun.star.beans.PropertyValue();
        lProperties[0].Name = "Hidden";
        lProperties[0].Value = true;
        XComponent xComponent = this.componentLoader.loadComponentFromURL(path, "_blank", 0, lProperties);

        this.document = (XTextDocument) UnoRuntime.queryInterface(XTextDocument.class, xComponent);
    }
}
