package dev.evt.openoffice;

/**
 * <h1>DocumentProperty</h1>
 * <p>
 * A document's property
 * </p>
 *
 * @author Eelke van Turnhout
 */
final public class DocumentProperty {
    private String name;
    private String value;

    /**
     * Constructs a DocumentProperty object.
     *
     * @param name  the name of the property
     * @param value the value of the property
     */
    public DocumentProperty(String name, String value) {
        this.name = name;
        this.value = value;
    }

    public String getName() {
        return this.name;
    }

    public String getValue() {
        return this.value;
    }
}
