package dev.evt.openoffice;

import java.util.ArrayList;

/**
 * Create a list of properties that determine how a document is saved
 * 
 * @author Eelke van Turnhout <eelketurnhout3@gmail.com>
 * @version 1.0
 */
public class DocumentProperties
{
    public final static String PROPERTY_OVERWRITE = "Overwrite";
    public final static String PROPERTY_FORMAT = "FilterName";
    public final static String PROPERTIES_AVAILABLE = PROPERTY_OVERWRITE + ", " + PROPERTY_FORMAT;

    public final static String VALUE_FORMAT_HTML = "HTML (StarWriter)";
    public final static String VALUE_FORMAT_DOC = "MS Word 97";
    public final static String VALUE_BOOLEAN_TRUE = String.valueOf(true);

    public final static String VALUES_AVAILABLE = VALUE_FORMAT_HTML + ", " + VALUE_BOOLEAN_TRUE;

    protected ArrayList<DocumentProperty> properties;

    /**
     * Constructs a DocumentProperties object
     */
    public DocumentProperties()
    {
        this.properties = new ArrayList<DocumentProperty>();
    }

    /**
     * Add a new property to the properties
     * 
     * @param property
     *            a new DocumentProperty
     * @throws Exception
     */
    public void addProperty(DocumentProperty property) throws Exception
    {
        this.validate(property);
        this.properties.add(property);
    }

    /**
     * Remove a property from the list
     * 
     * @param property
     *            the DocumentProperty that needs removing
     */
    public void removeProperty(DocumentProperty property)
    {
        this.properties.remove(property);
    }

    /**
     * Get the list of properties
     * 
     * @return an array list of DocumentProperty objects
     */
    public ArrayList<DocumentProperty> getProperties()
    {
        return this.properties;
    }

    /**
     * Get the amount of properties currently int the list
     * 
     * @return list size as integer
     */
    public int count()
    {
        return this.properties.size();
    }

    /**
     * Validates the property by checking the property name and value against
     * the PROPERTIES_AVAILABLE and VALUES_AVAILABLE constants
     * 
     * @param property
     *            The DocumentProperty that needs validating
     * @throws Exception
     */
    protected void validate(DocumentProperty property) throws Exception
    {
        if (!PROPERTIES_AVAILABLE.contains(property.getName())) {
            throw new Exception("Invalid property " + property.getName() + ", the following options are available: "
                    + PROPERTIES_AVAILABLE + ".");
        }

        if (!VALUES_AVAILABLE.contains(property.getValue())) {
            throw new Exception("Invalid value " + property.getValue() + ", the following options are available: "
                    + VALUES_AVAILABLE + ".");
        }
    }
}
