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

        TextDocument document = (TextDocument) manager.openNew(DocumentManager.NEW_TEXT_DOCUMENT);
        assertThat(document, instanceOf(TextDocument.class));

        document.setName(name);

        assertEquals("testing", document.getName());
        assertEquals(TextDocument.EXTENSION_DEFAULT, document.getExtension());

        manager.save(document);
        manager.close(document);

        document = (TextDocument) manager.open(name + TextDocument.EXTENSION_DEFAULT);

        DocumentProperties properties = new DocumentProperties();
        properties.addProperty(new DocumentProperty(TextDocument.PROPERTY_FORMAT, TextDocument.PROPERTY_VALUE_FORMAT_HTML));
        document.setProperties(properties);
        document.setExtension(TextDocument.EXTENSION_HTML);

        manager.save(document);
        manager.close(document);

        document = (TextDocument) manager.open(name + TextDocument.EXTENSION_HTML);

        assertEquals(name, document.getName());
        assertEquals(TextDocument.EXTENSION_HTML, document.getExtension());

        manager.close(document);
        manager.delete(document);
        document.setExtension(TextDocument.EXTENSION_DEFAULT);
        manager.delete(document);
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
