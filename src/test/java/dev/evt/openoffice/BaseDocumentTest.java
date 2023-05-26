package dev.evt.openoffice;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import org.junit.jupiter.api.Test;

public class BaseDocumentTest
{
    protected Connection connection;

    public BaseDocumentTest() throws Exception
    {
        QuickConnect quickConnect = new QuickConnect();
        this.connection = quickConnect.getConnection();
    }

    /**
     * Test if the properties are set correctly and accessible
     */
    @Test
    public void testProperties()
    {
        String folder = "/some/location/";
        String file = "somefile.docx";
        String name = "somefile";
        String extension = ".docx";
        BaseDocument baseDocument = new BaseDocument(this.connection, folder, file, null);

        assertEquals(name, baseDocument.getName());
        assertEquals(extension, baseDocument.getExtension());
    }

    /**
     * Test if the file can be passed as null
     */
    @SuppressWarnings("unused")
    @Test
    public void testFileIsNull()
    {
        String folder = "/validfolder/";
        BaseDocument baseDocument = new BaseDocument(this.connection, folder, null, null);
    }

    /**
     * Check if an IllegalArgumentException is thrown when the folder doesn't
     * end with a forward slash
     */
    @SuppressWarnings("unused")
    @Test
    public void testInvalidFolder()
    {
        assertThrows(IllegalArgumentException.class, () -> {
            String folder = "invalidfolder";
            String file = "somefile.doc";
            BaseDocument baseDocument = new BaseDocument(this.connection, folder, file, null);
        });
    }

    /**
     * Test if an IllegalArgumentException is thrown when supplying a file
     * without it's extension
     */
    @SuppressWarnings("unused")
    @Test
    public void testInvalidFile()
    {
        assertThrows(IllegalArgumentException.class, () -> {
            String folder = "validfolder/";
            String file = "somefile";
            BaseDocument baseDocument = new BaseDocument(this.connection, folder, file, null);
        });
    }
}
