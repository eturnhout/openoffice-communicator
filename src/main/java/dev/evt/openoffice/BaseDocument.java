package dev.evt.openoffice;

/**
 * <h1>BaseDocument</h1>
 * <p>
 * A base document class used for extending.
 * </p>
 * 
 * @author Eelke van Turnhout <eelketurnhout3@gmail.com>
 * @version 1.0
 */
public class BaseDocument extends BaseConnection
{
    public final static String PROPERTY_OVERWRITE = "Overwrite";
    public final static String PROPERTY_FORMAT = "FilterName";
    public final static String PROPERTIES_AVAILABLE = PROPERTY_OVERWRITE + ", " + PROPERTY_FORMAT;

    public final static String VALUE_BOOLEAN_TRUE = String.valueOf(true);

    /**
     * Full path to the directory of the document.
     */
    protected String folder;

    /**
     * Name of this document.
     */
    protected String name;

    /**
     * The document's extension.
     */
    protected String extension;

    /**
     * The document's properties (important when saving).
     */
    protected DocumentProperties properties;

    /**
     * Construct a BaseDocument object.
     * 
     * @param connection
     *            Connection instance.
     * @param folder
     *            Full path to the directory of the document.
     * @param file
     *            (optional) File name plus extension.
     * @param properties
     *            (optional) The document's properties.
     */
    public BaseDocument(Connection connection, String folder, String file, DocumentProperties properties)
    {
        super(connection);
        this.setFolder(folder);
        if (file != null) {
            this.parseFile(file);
        }
        this.properties = properties;
    }

    /**
     * Set the folder location of the document.
     * 
     * @param folder
     *            The location of the document.
     */
    public void setFolder(String folder) throws IllegalArgumentException
    {
        if (folder.lastIndexOf("/") != folder.length() - 1) {
            throw new IllegalArgumentException("Invalid folder \"" + folder + "\" Make sure the folder ends with a \"/\".");
        }

        this.folder = folder;
    }

    /**
     * Get the folder location of the document.
     * 
     * @return The location of the document.
     */
    public String getFolder()
    {
        return this.folder;
    }

    /**
     * Set the document's name.
     * 
     * @param name
     *            Name of the document.
     */
    public void setName(String name)
    {
        this.name = name;
    }

    /**
     * Get the document's name.
     * 
     * @return the document's name.
     */
    public String getName()
    {
        return this.name;
    }

    /**
     * The document's extension. This is useful when saving the document in
     * another format.
     * 
     * @param extension
     *            The document's extension.
     */
    public void setExtension(String extension)
    {
        this.extension = extension;
    }

    /**
     * Get the document's extension.
     * 
     * @return the document's extension.
     */
    public String getExtension()
    {
        return this.extension;
    }

    /**
     * Set the document's properties.
     * 
     * @param properties
     *            A DocumentProperties object.
     */
    public void setProperties(DocumentProperties properties)
    {
        this.properties = properties;
    }

    /**
     * Get the document's properties.
     * 
     * @return the document's properties.
     */
    public DocumentProperties getProperties()
    {
        return this.properties;
    }

    /**
     * Splits the file and sets the fileName and fileExtension properties.
     * 
     * @param file
     *            File name + extension.
     */
    protected void parseFile(String file) throws IllegalArgumentException
    {
        if (file.length() < 3) {
            throw new IllegalArgumentException("Invalid file \"" + file + "\" provided.");
        } else if (!file.contains(".")) {
            throw new IllegalArgumentException("Invalid file \"" + file + "\" no extension found.");
        }

        String name = file.substring(0, file.lastIndexOf("."));
        String extension = file.substring(file.lastIndexOf("."), file.length());
        this.setName(name);
        this.setExtension(extension);
    }
}
