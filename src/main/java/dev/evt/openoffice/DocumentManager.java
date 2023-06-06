package dev.evt.openoffice;

import java.io.InputStream;
import java.io.UnsupportedEncodingException;

import com.sun.star.beans.PropertyValue;
import com.sun.star.frame.XStorable;
import com.sun.star.io.IOException;
import com.sun.star.io.XInputStream;
import com.sun.star.io.XOutputStream;
import com.sun.star.lang.XComponent;
import com.sun.star.ucb.CommandAbortedException;
import com.sun.star.ucb.XSimpleFileAccess;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.util.CloseVetoException;
import com.sun.star.util.XCloseable;

/**
 * <h1>DocumentManager</h1>
 * <p>
 * Manager that handles document storage.
 * </p>
 *
 * @author Eelke van Turnhout
 */
public class DocumentManager {
    public final static int NEW_TEXT_DOCUMENT = 1;

    protected String folder;
    protected Connection connection;
    protected XSimpleFileAccess fileAccess;

    /**
     * Constructs a DocumentManager object.
     *
     * @param connection
     *                   The connection to a running OpenOffice service.
     * @param folder
     *                   The storage folder on the filesystem for the documents.
     * @throws Exception
     */
    public DocumentManager(Connection connection, String folder) throws Exception {
        this.setFolder(folder);

        this.connection = connection;
        this.fileAccess = (XSimpleFileAccess) UnoRuntime.queryInterface(XSimpleFileAccess.class,
                connection.getServiceFactory().createInstance("com.sun.star.ucb.SimpleFileAccess"));
    }

    /**
     * Set the location where the documents will be handled.
     *
     * @param folder
     *               Full system path to the folder.
     */
    public void setFolder(String folder) {
        this.folder = folder;
    }

    /**
     * Set the location where the documents will be handled.
     *
     * @return folder The location of the documents storage folder.
     */
    public String getFolder() {
        return this.folder;
    }

    /**
     * Creates the file from a input stream.
     *
     * @param file
     *                    The file name plus extension.
     * @param inputStream
     *                    A file's input stream.
     * @throws CommandAbortedException
     * @throws Exception
     * @throws java.io.IOException
     */
    @SuppressWarnings("unused")
    public void createFromInputStream(String file, InputStream inputStream)
            throws CommandAbortedException, Exception, java.io.IOException {
        if (this.exists(file)) {
            throw new Exception("There's already a file called " + file + ".");
        }

        String fullPath = this.folder + file;
        XOutputStream outputStream = (XOutputStream) this.fileAccess.openFileReadWrite(fullPath).getInputStream();

        byte[] bytes = new byte[128];
        int read = 0;

        while ((read = inputStream.read(bytes)) != -1) {
            outputStream.writeBytes(bytes);
        }

        outputStream.flush();
        outputStream.closeOutput();
        inputStream.close();
    }

    /**
     * Get the document's raw content as a string. Handy for html documents.
     *
     * @param document
     *                 The document from which to get the content. This document
     *                 must
     *                 be saved and available on the filesystem.
     * @return the document's content in string format.
     * @throws CommandAbortedException
     * @throws Exception
     * @throws UnsupportedEncodingException
     */
    @SuppressWarnings("unused")
    public String getRawContent(BaseDocument document)
            throws CommandAbortedException, Exception, UnsupportedEncodingException {
        if (!this.exists(document.getName() + document.getExtension())) {
            throw new Exception("No file called " + document.getName() + document.getExtension() + " found.");
        }

        String fullPath = document.getFolder() + document.getName() + document.getExtension();

        XInputStream inputStream = (XInputStream) this.fileAccess.openFileReadWrite(fullPath).getInputStream();
        StringBuffer buffer = new StringBuffer();
        byte[][] bytes = new byte[1024][];
        int read = 0;

        while ((read = inputStream.readSomeBytes(bytes, 1024)) != 0) {
            buffer.append(new String(bytes[0], "UTF-8"));
        }

        inputStream.closeInput();

        return buffer.toString();
    }

    /**
     * Open a new document.
     *
     * @param type
     *             The type of document. See the managers static properties for
     *             the available options.
     * @return a new document, the kind depends on the type argument or null on
     *         failure.
     * @throws java.lang.Exception
     */
    public BaseDocument openNew(int type) throws java.lang.Exception {
        if (type == NEW_TEXT_DOCUMENT) {
            return new TextDocument(this.connection, this.folder, null, null);
        }

        return null;
    }

    /**
     * Open an existing document.
     *
     * @param file
     *             The full file name plus extension.
     * @return a TextDocument object
     * @throws java.lang.Exception
     */
    public BaseDocument open(String file) throws java.lang.Exception {
        BaseDocument baseDocument = new BaseDocument(this.connection, this.folder, file, null);

        if (!this.exists(baseDocument.getName() + baseDocument.getExtension())) {
            throw new Exception(
                    "File with file name " + baseDocument.getName() + baseDocument.getExtension() + " does not exist.");
        }

        if (TextDocument.isValidExtension(baseDocument.getExtension())) {
            return new TextDocument(this.connection, this.folder, file, null);
        }

        return null;
    }

    /**
     * Saves a document.
     *
     * @param document
     *                 The BaseDocument object that needs saving.
     * @throws IOException
     */
    public void save(BaseDocument document) throws IOException {
        XStorable storeable = null;

        if (document instanceof TextDocument) {
            TextDocument doc = (TextDocument) document;
            storeable = (XStorable) UnoRuntime.queryInterface(XStorable.class, doc.getDocument());
        }

        int indexes = document.getProperties().count();
        PropertyValue[] options = new PropertyValue[indexes];
        int index = 0;

        for (DocumentProperty property : document.getProperties().getProperties()) {
            PropertyValue option = new PropertyValue();
            option.Name = property.getName();
            option.Value = (property.getValue() == BaseDocument.VALUE_BOOLEAN_TRUE) ? true : property.getValue();

            options[index] = option;
            index++;
        }

        String file = document.getName() + document.getExtension();
        String fullPath = "file://" + this.folder + file;

        storeable.storeToURL(fullPath, options);
    }

    /**
     * Delete a document. This happens based on the document's name and
     * extension.
     *
     * @param document
     *                 The document that needs deleting.
     * @throws CommandAbortedException
     * @throws Exception
     */
    public void delete(BaseDocument document) throws CommandAbortedException, Exception {
        String filePath = this.folder + document.getName() + document.getExtension();
        this.fileAccess.kill(filePath);
    }

    /**
     * Closes the document.
     *
     * @param document
     *                 The BaseDocument object that needs closing.
     * @throws CloseVetoException
     */
    public void close(BaseDocument document) throws CloseVetoException {
        XStorable storable = null;

        if (document instanceof TextDocument) {
            TextDocument doc = (TextDocument) document;
            storable = (XStorable) UnoRuntime.queryInterface(XStorable.class, doc.getDocument());
        }

        XCloseable closeable = (XCloseable) UnoRuntime.queryInterface(XCloseable.class, storable);

        if (closeable != null) {
            closeable.close(false);
        } else {
            XComponent component = (XComponent) UnoRuntime.queryInterface(XComponent.class, storable);
            component.dispose();
        }
    }

    /**
     * Check if a file already exists.
     *
     * @param file
     *             File name plus extension.
     * @return true if the file exists, returns false otherwise.
     * @throws CommandAbortedException
     * @throws Exception
     */
    public boolean exists(String file) throws CommandAbortedException, Exception {
        String filePath = this.folder + file;

        return this.fileAccess.exists(filePath);
    }
}
