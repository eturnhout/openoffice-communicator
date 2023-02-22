package dev.evt.openoffice;

import static org.hamcrest.CoreMatchers.instanceOf;
import static org.hamcrest.MatcherAssert.assertThat;
import static org.junit.Assert.assertEquals;

import org.junit.Test;

import com.sun.star.io.IOException;
import com.sun.star.lang.IllegalArgumentException;

final public class DocumentManagerTest {
    private Connection connection;
    private ConfigReader configReader;

    public DocumentManagerTest() throws Exception {
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
    public void testTextDocumentToHtml() throws Exception {
        this.convertAndDelete(TextDocument.EXTENSION_HTML, TextDocument.PROPERTY_VALUE_FORMAT_HTML);
    }

    /**
     * Test if a new TextDocument can be opened, saved, closed, and deleted
     *
     * @throws IllegalArgumentException
     * @throws IOException
     */
    @Test
    public void testTextDocumentPdfExport() throws Exception {
        this.convertAndDelete(TextDocument.EXTENSION_PDF, TextDocument.PROPERTY_VALUE_FORMAT_PDF);
    }

    private void convertAndDelete(String extension, String format) throws Exception {
        String folder = this.configReader.getConfig("folder");
        String name = "testing";
        DocumentManager manager = new DocumentManager(this.connection, folder);

        // Create a default document to convert
        TextDocument defaultDocument = (TextDocument) manager.openNew(DocumentManager.NEW_TEXT_DOCUMENT);
        defaultDocument.setName(name);

        // We need to set the property "Overwrite" so storeToURL won't fail
        defaultDocument.getProperties()
                .addProperty(new DocumentProperty(TextDocument.PROPERTY_OVERWRITE, BaseDocument.VALUE_BOOLEAN_TRUE));

        manager.save(defaultDocument);

        assertThat(defaultDocument, instanceOf(TextDocument.class));
        assertEquals(name, defaultDocument.getName());
        assertEquals(TextDocument.EXTENSION_DEFAULT, defaultDocument.getExtension());

        manager.close(defaultDocument);

        // Save default document as html
        TextDocument newDocument = (TextDocument) manager.open(name + TextDocument.EXTENSION_DEFAULT);

        newDocument.getProperties().addProperty(new DocumentProperty(TextDocument.PROPERTY_FORMAT, format));
        newDocument.setExtension(extension);

        manager.save(newDocument);

        assertThat(newDocument, instanceOf(TextDocument.class));
        assertEquals(name, newDocument.getName());
        assertEquals(extension, newDocument.getExtension());

        manager.close(newDocument);
        manager.delete(newDocument);
        manager.delete(defaultDocument);
    }
}
