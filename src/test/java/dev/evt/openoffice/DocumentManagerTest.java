package dev.evt.openoffice;

import static org.hamcrest.CoreMatchers.instanceOf;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertThat;

import org.junit.Test;

import com.sun.star.io.IOException;
import com.sun.star.lang.IllegalArgumentException;

public class DocumentManagerTest
{
    protected Connection connection;
    protected ConfigReader configReader;

    public DocumentManagerTest() throws Exception
    {
        QuickConnect quickConnect = new QuickConnect();
        this.connection = quickConnect.getConnection();
        this.configReader = quickConnect.getConfigReader();
    }

    /**
     * Test if a new TextDocument can be opened, saved, closed, and deleted
     *
     * @throws IllegalArgumentException
     * @throws IOException
     */
    @Test
    public void testTextDocumentToHtml() throws Exception
    {
        String folder = this.configReader.getConfig("folder");
        String name = "testing";
        DocumentManager manager = new DocumentManager(this.connection, folder);

        // Create a default document to convert
        TextDocument defaultDocument = (TextDocument) manager.openNew(DocumentManager.NEW_TEXT_DOCUMENT);
        defaultDocument.setName(name);

        // We need to set the property "Overwrite" so storeToURL won't fail
        DocumentProperties properties = new DocumentProperties();
        properties.addProperty(new DocumentProperty(TextDocument.PROPERTY_OVERWRITE, BaseDocument.VALUE_BOOLEAN_TRUE));
        defaultDocument.setProperties(properties);

        manager.save(defaultDocument);

        assertThat(defaultDocument, instanceOf(TextDocument.class));
        assertEquals(name, defaultDocument.getName());
        assertEquals(TextDocument.EXTENSION_DEFAULT, defaultDocument.getExtension());

        manager.close(defaultDocument);

        // Save default document as html
        TextDocument htmlDocument = (TextDocument) manager.open(name + TextDocument.EXTENSION_DEFAULT);

        // We can re-use the properties from the default document and add the html format prop
        properties.addProperty(new DocumentProperty(TextDocument.PROPERTY_FORMAT, TextDocument.PROPERTY_VALUE_FORMAT_HTML));
        htmlDocument.setProperties(properties);
        htmlDocument.setExtension(TextDocument.EXTENSION_HTML);

        manager.save(htmlDocument);

        assertThat(htmlDocument, instanceOf(TextDocument.class));
        assertEquals(name, htmlDocument.getName());
        assertEquals(TextDocument.EXTENSION_HTML, htmlDocument.getExtension());

        manager.close(htmlDocument);
        manager.delete(htmlDocument);
        manager.delete(defaultDocument);
    }

    /**
     * Test if a new TextDocument can be opened, saved, closed, and deleted
     *
     * @throws IllegalArgumentException
     * @throws IOException
     */
    @Test
    public void testTextDocumentPdfExport() throws Exception
    {
        String folder = this.configReader.getConfig("folder");
        String name = "testing";
        DocumentManager manager = new DocumentManager(this.connection, folder);

        TextDocument document = (TextDocument) manager.openNew(DocumentManager.NEW_TEXT_DOCUMENT);
        assertThat(document, instanceOf(TextDocument.class));

        document.setName(name);

        assertEquals("testing", document.getName());
        assertEquals(TextDocument.EXTENSION_DEFAULT, document.getExtension());

        manager.save(document);
        manager.close(document);

        document = (TextDocument) manager.open(name + TextDocument.EXTENSION_DEFAULT);
        document.setExtension(TextDocument.EXTENSION_PDF);
        DocumentProperties properties = new DocumentProperties();
        properties.addProperty(new DocumentProperty(TextDocument.PROPERTY_FORMAT, TextDocument.PROPERTY_VALUE_FORMAT_PDF));
        document.setProperties(properties);

        manager.save(document);
        manager.close(document);
        manager.delete(document);
        document.setExtension(TextDocument.EXTENSION_DEFAULT);
        manager.delete(document);
    }
}
