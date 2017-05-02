package dev.evt.openoffice;

import java.io.InputStream;

import com.sun.star.beans.PropertyValue;
import com.sun.star.frame.XStorable;
import com.sun.star.io.IOException;
import com.sun.star.io.XInputStream;
import com.sun.star.io.XOutputStream;
import com.sun.star.lang.XComponent;
import com.sun.star.ucb.CommandAbortedException;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.util.CloseVetoException;
import com.sun.star.util.XCloseable;

/**
 * Manager that handles document storage
 *
 * @author Eelke van Turnhout <eelketurnhout3@gmail.com>
 * @version 1.0
 */
public class DocumentManager extends BaseConnection
{
    /**
     * List of comma separated available extensions
     */
    public final static String AVAILABLE_EXTENSIONS = ".doc, .docx .html";

    protected String folder;
    protected FileAccess fileAccess;

    /**
     * Constructs a DocumentManager object
     * 
     * @param connection The connection to a running OpenOffice service
     * @param folder The storage folder on the filesystem for the documents
     * @throws Exception
     */
    public DocumentManager(Connection connection, String folder) throws Exception
    {
        super(connection);
        this.setFolder(folder);

        try {
            this.fileAccess = new FileAccess(connection);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * Set the location where the documents will be handled
     *
     * @param folder
     */
    public void setFolder(String folder)
    {
        this.folder = folder;
    }
    
    /**
     * Set the location where the documents will be handled
     * 
     * @return folder The location of the documents storage folder
     */
    public String getFolder()
    {
        return this.folder;
    }
    
    /**
     * Creates the file from a input stream
     * 
     * @param fileName the file's name plus extension
     * @param fileExtension the file's extension
     * @param inputStream the input stream 
     * @throws CommandAbortedException
     * @throws Exception
     * @throws java.io.IOException
     */
    public void createFromInputStream(String fileName, String fileExtension, InputStream inputStream) throws CommandAbortedException, Exception, java.io.IOException
    {
        if (this.exists(fileName, fileExtension)) {
            throw new Exception("There's already a file called " + fileName + fileExtension + ".");
        }

        String fullPath = this.folder + fileName + fileExtension;
        XOutputStream outputStream = (XOutputStream) fileAccess.getStream(fullPath).getOutputStream();

        byte[] bytes = new byte[128];
        int read = 0;
        int amount = 0;

        while ((read = inputStream.read(bytes)) != -1) {
            amount += bytes.length;
            outputStream.writeBytes(bytes);
        }

        System.out.println("Read: " + amount + " bytes");
        
        outputStream.flush();
        outputStream.closeOutput();
        inputStream.close();
    }

    /**
     * Get the document's raw content as a string. Handy for html documents.
     *
     * @param document The TextDocument from which to get the content. This document must be stored and available on the filesystem.
     * @return the document's content in string format
     * @throws CommandAbortedException
     * @throws Exception
     */
    public String getRawContent(TextDocument document) throws CommandAbortedException, Exception
    {
        if (!this.exists(document.getFileName(), document.getFileExtension())) {
            throw new Exception("No file called " + document.getFileName() + document.getFileExtension() + " found.");
        }

        String fullPath = document.getFolder() + document.getFileName() + document.getFileExtension();

        XInputStream inputStream = (XInputStream) this.fileAccess.getStream(fullPath).getInputStream();
        StringBuffer buffer = new StringBuffer();
        byte[][] bytes = new byte[1024][];
        int read = 0;

        while ((read = inputStream.readSomeBytes(bytes, 1024)) != 0) {
            buffer.append(new String(bytes[0]));
        }

        inputStream.closeInput();

        return buffer.toString();
    }
    
    
    /**
     * Open an existing document
     * 
     * @param fileName The document's name
     * @param fileExtension The document's extension
     * @return a TextDocument object
     * @throws CommandAbortedException
     * @throws Exception
     */
    public TextDocument open(String fileName, String fileExtension) throws CommandAbortedException, Exception
    {
        if (!AVAILABLE_EXTENSIONS.contains(fileExtension)) {
            throw new Exception("Unsupported file extension \"" + fileExtension + "\", does not match any of the following available extensions: " + AVAILABLE_EXTENSIONS + ".");
        }

        if (!this.exists(fileName, fileExtension)) {
            throw new Exception("File with file name " + fileName + fileExtension + " does not exist.");
        }

        return new TextDocument(this.connection, this.folder, fileName, fileExtension);
    }

    /**
     * Saves a document
     * 
     * @param document The TextDocument object than needs saving
     * @param properties Properties that determine how the document will be saved 
     * @throws IOException
     */
    public void save(TextDocument document, DocumentProperties properties) throws IOException
    {
        XStorable storeable = (XStorable) UnoRuntime.queryInterface(XStorable.class, document.getDocument());
        int indexes = properties.count();
        PropertyValue[] options = new PropertyValue[indexes];
        int index = 0;

        for (DocumentProperty property : properties.getProperties()) {
            PropertyValue option = new PropertyValue();
            option.Name = property.getName();
            option.Value = (property.getValue() == DocumentProperties.VALUE_BOOLEAN_TRUE) ? true : property.getValue();

            options[index] = option;
            index++;
        }

        String fileName = document.getFileName() + document.getFileExtension();
        String fullPath = "file://" + this.folder + fileName;

        storeable.storeAsURL(fullPath, options);
    }
    
    /**
     * Delete a document
     * 
     * @param document The document that needs deleting
     * @throws CommandAbortedException
     * @throws Exception
     */
    public void delete(TextDocument document) throws CommandAbortedException, Exception
    {
        String filePath = this.folder + document.getFileName() + document.getFileExtension();
        this.fileAccess.delete(filePath);
    }

    /**
     * Check if a file already exists
     *
     * @param fileName Name of the file
     * @param fileExtension The file's extension
     * @return true if the file exists, returns false otherwise
     * @throws CommandAbortedException
     * @throws Exception
     */
    public boolean exists(String fileName, String fileExtension) throws CommandAbortedException, Exception
    {
        String filePath = this.folder + fileName + fileExtension;

        if (!this.fileAccess.exists(filePath)) {
            return false;
        }

        return true;
    }

    /**
     * 
     * @param document The TextDocument object that needs closing
     * @throws CloseVetoException
     */
    public void close(TextDocument document) throws CloseVetoException
    {
        XStorable storeable = (XStorable) UnoRuntime.queryInterface(XStorable.class, document.getDocument());
        XCloseable closeable = (XCloseable) UnoRuntime.queryInterface(XCloseable.class, storeable);

        if (closeable != null) {
            closeable.close(false);
        } else {
            XComponent component = (XComponent)UnoRuntime.queryInterface(XComponent.class, storeable);
            component.dispose();
        }
    }
}
