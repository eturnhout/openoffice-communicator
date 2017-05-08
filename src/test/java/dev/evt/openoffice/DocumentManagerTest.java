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
    public void testTextDocument() throws Exception
    {
        String folder = this.configReader.getConfig("folder");
        String name = "testing";
        DocumentManager manager = new DocumentManager(this.connection, folder);

        TextDocument document = (TextDocument) manager.openNew(DocumentManager.NEW_TEXT_DOCUMENT);
        assertThat(document, instanceOf(TextDocument.class));

        document.setName(name);

        assertEquals("testing", document.getName());
        assertEquals(".doc", document.getExtension());

        manager.save(document);
        manager.close(document);

        document = (TextDocument) manager.open(name + ".doc");

        DocumentProperties properties = new DocumentProperties();
        properties.addProperty(new DocumentProperty(TextDocument.PROPERTY_FORMAT, TextDocument.PROPERTY_VALUE_FORMAT_HTML));
        document.setProperties(properties);
        document.setExtension(".html");

        manager.save(document);
        manager.close(document);

        document = (TextDocument) manager.open(name + ".html");

        assertEquals(name, document.getName());
        assertEquals(".html", document.getExtension());

        manager.close(document);
        manager.delete(document);
        document.setExtension(".doc");
        manager.delete(document);
    }
}
