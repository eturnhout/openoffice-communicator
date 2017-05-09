package dev.evt.openoffice;

import com.sun.star.io.XStream;
import com.sun.star.ucb.CommandAbortedException;
import com.sun.star.ucb.XSimpleFileAccess;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;

/**
 * <h1>FileAccess</h1>
 * <p>
 * Simple file interaction
 * </p>
 * 
 * @author Eelke van Turnhout <eelketurnhout3@gmail.com>
 * @version 1.0
 */
class FileAccess extends BaseConnection
{
    protected XSimpleFileAccess fileAccess;

    /**
     * Construct a FileAccess object
     * 
     * @param Connection
     *            A connection to OpenOffice.
     * @throws Exception
     */
    public FileAccess(Connection connection) throws Exception
    {
        super(connection);
        this.initialize();
    }

    /**
     * Check if a file exists.
     * 
     * @param filePath
     *            Full path to the file.
     * @return true or false.
     * @throws CommandAbortedException
     * @throws com.sun.star.uno.Exception
     */
    public boolean exists(String filePath) throws CommandAbortedException, com.sun.star.uno.Exception
    {
        return this.fileAccess.exists(filePath);
    }

    /**
     * Delete a file.
     * 
     * @param filePath
     *            The full path to the file.
     * @throws CommandAbortedException
     * @throws Exception
     */
    public void delete(String filePath) throws CommandAbortedException, Exception
    {
        this.fileAccess.kill(filePath);
    }

    /**
     * General file stream to access the input or output stream of a file.
     * 
     * @param path
     *            Full path to the file.
     * @return General stream to access the input/output streams
     * @throws CommandAbortedException
     * @throws Exception
     */
    public XStream getStream(String path) throws CommandAbortedException, Exception
    {
        return this.fileAccess.openFileReadWrite(path);
    }

    /**
     * Initializes the object XSimpleFileAccess object needed to interact with
     * the filesystem.
     * 
     * @throws Exception
     */
    private void initialize() throws Exception
    {
        Object fileAccessObject = this.connection.getServiceFactory().createInstance("com.sun.star.ucb.SimpleFileAccess");
        this.fileAccess = (XSimpleFileAccess) UnoRuntime.queryInterface(XSimpleFileAccess.class, fileAccessObject);
    }
}
