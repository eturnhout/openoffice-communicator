package dev.evt.openoffice;

import java.io.BufferedReader;
import java.io.Reader;
import java.util.Scanner;

import com.sun.star.beans.PropertyValue;
import com.sun.star.io.IOException;
import com.sun.star.io.XInputStream;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.XComponent;
import com.sun.star.text.XText;
import com.sun.star.text.XTextDocument;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;

/**
 * Used to interact with text documents.
 * 
 * @author Eelke van Turnhout <eelketurnhout3@gmail.com>
 * @version 1.0
 */
public class TextDocument extends Desktop
{
    final static String NEW_DOCUMENT = "private:factory/swriter";

    protected String folder;
    protected String fileName;
    protected String fileExtension;
    protected XTextDocument document;

    public TextDocument(Connection connection, String folder, String fileName, String fileExtension)
            throws IOException, IllegalArgumentException
    {
        super(connection);
        this.setFolder(folder);
        this.setFileName(fileName);
        this.setFileExtension(fileExtension);
        this.initialize();
    }

    /**
     * Set the folder location of the document
     * 
     * @param folder
     *            The location of the document
     */
    public void setFolder(String folder)
    {
        this.folder = folder;
    }

    /**
     * Get the folder location of the document
     * 
     * @return The location of the document
     */
    public String getFolder()
    {
        return this.folder;
    }

    /**
     * Set the document's name
     * 
     * @param fileName
     *            Name of the document
     */
    public void setFileName(String fileName)
    {
        this.fileName = fileName;
    }

    /**
     * Get the document's name.
     * 
     * @return the document's name
     */
    public String getFileName()
    {
        return this.fileName;
    }

    /**
     * The document's extension. This is useful when saving the document in
     * another format.
     * 
     * @param fileExtension
     *            The document's extension
     */
    public void setFileExtension(String fileExtension)
    {
        this.fileExtension = fileExtension;
    }

    /**
     * Get the document's extension.
     * 
     * @return the document's extension
     */
    public String getFileExtension()
    {
        return this.fileExtension;
    }

    /**
     * Get the documents text
     * 
     * @return the documents text in String format
     */
    public String getText()
    {
        XText text = this.document.getText();
        return text.getString();
    }

    /**
     * Get the XTextDocument object. This can be used to interact with the
     * document.
     * 
     * @return the XTextDocument object for this document
     */
    public XTextDocument getDocument()
    {
        return this.document;
    }

    /**
     * Initialize this object. This will fill in the XTextDocument object with a
     * new or existing file if it already exists
     *
     * @throws IOException
     * @throws IllegalArgumentException
     */
    private void initialize() throws IOException, IllegalArgumentException
    {
        String path = TextDocument.NEW_DOCUMENT;

        if (this.fileName != null) {
            path = "file://" + this.folder + this.fileName + this.fileExtension;
        }

        PropertyValue[] lProperties = new PropertyValue[1];
        lProperties[0] = new com.sun.star.beans.PropertyValue();
        lProperties[0].Name = "Hidden";
        lProperties[0].Value = true;
        XComponent xComponent = this.componentLoader.loadComponentFromURL(path, "_blank", 0, lProperties);

        this.document = (XTextDocument) UnoRuntime.queryInterface(XTextDocument.class, xComponent);
    }
}
