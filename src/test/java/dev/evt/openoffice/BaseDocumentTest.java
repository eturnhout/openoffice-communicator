package dev.evt.openoffice;

import static org.junit.Assert.assertEquals;

import org.junit.Test;

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
        BaseDocument baseDocument = new BaseDocument(this.connection, folder, file);

        assertEquals(name, baseDocument.getName());
        assertEquals(extension, baseDocument.getExtension());
    }

    /**
     * Check if an IllegalArgumentException is thrown when the folder doesn't
     * end with a forward slash
     */
    @SuppressWarnings("unused")
    @Test(expected = IllegalArgumentException.class)
    public void testInvalidFolder()
    {
        String folder = "invalidfolder";
        String file = "somefile.doc";
        BaseDocument baseDocument = new BaseDocument(this.connection, folder, file);
    }

    /**
     * Test if an IllegalArgumentException is thrown when supplying a file
     * without it's extension
     */
    @SuppressWarnings("unused")
    @Test(expected = IllegalArgumentException.class)
    public void testInvalidFile()
    {
        String folder = "validfolder/";
        String file = "somefile";
        BaseDocument baseDocument = new BaseDocument(this.connection, folder, file);
    }
}
