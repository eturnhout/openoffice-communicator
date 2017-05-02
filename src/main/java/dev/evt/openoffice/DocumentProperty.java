package dev.evt.openoffice;

/**
 * A document's property 
 *
 * @author Eelke van Turnhout <eelketurnhout3@gmail.com>
 * @version 1.0
 */
public class DocumentProperty
{
    protected String name;
    protected String value;
    
    /**
     * Constructs a DocumentProperty object
     *
     * @param name The name of the property
     * @param value The value of the property
     */
    public DocumentProperty(String name, String value)
    {
        this.setName(name);
        this.setValue(value);
    }

    public void setName(String name)
    {
        this.name = name;
    }

    public String getName()
    {
        return this.name;
    }

    public void setValue(String value)
    {
        this.value = value;
    }

    public String getValue()
    {
        return this.value;
    }
}
