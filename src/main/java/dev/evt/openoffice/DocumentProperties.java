package dev.evt.openoffice;

import java.util.ArrayList;

/**
 * <h1>DocumentProperties</h1>
 * <p>
 * Create a list of properties that determine how a document is saved.
 * </p>
 * 
 * @author Eelke van Turnhout <eelketurnhout3@gmail.com>
 * @version 1.0
 */
public class DocumentProperties
{
    protected ArrayList<DocumentProperty> properties;

    /**
     * Constructs a DocumentProperties object.
     */
    public DocumentProperties()
    {
        this.properties = new ArrayList<DocumentProperty>();
    }

    /**
     * Add a new property to the properties.
     * 
     * @param property
     *            A new DocumentProperty object.
     * @throws Exception
     */
    public void addProperty(DocumentProperty property) throws Exception
    {
        this.validate(property);
        this.properties.add(property);
    }

    /**
     * Remove a property from the list.
     * 
     * @param property
     *            the DocumentProperty that needs removing.
     */
    public void removeProperty(DocumentProperty property)
    {
        this.properties.remove(property);
    }

    /**
     * Get the list of properties.
     * 
     * @return an array list of DocumentProperty objects.
     */
    public ArrayList<DocumentProperty> getProperties()
    {
        return this.properties;
    }

    /**
     * Get the amount of properties currently in the list.
     * 
     * @return list size as integer.
     */
    public int count()
    {
        return this.properties.size();
    }

    /**
     * Validates the property by checking the property name and value against
     * the PROPERTIES_AVAILABLE and VALUES_AVAILABLE constants in every document
     * class.
     * 
     * @param property
     *            The DocumentProperty that needs validating.
     * @throws Exception
     */
    protected void validate(DocumentProperty property) throws Exception
    {
        if (!TextDocument.PROPERTIES_AVAILABLE.contains(property.getName())) {
            throw new Exception("Invalid property " + property.getName() + ", the following options are available: " + TextDocument.PROPERTIES_AVAILABLE + ".");
        }
    }
}
