package dev.evt.openoffice;

import static org.hamcrest.CoreMatchers.instanceOf;
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
        DocumentManager manager = new DocumentManager(this.connection, folder);

        TextDocument document = (TextDocument) manager.open("easychair.docx");
        assertThat(document, instanceOf(TextDocument.class));

        document.setName("testing");

        manager.save(document);
        manager.close(document);
        manager.delete(document);
    }
}
