package dev.evt.openoffice;

import com.sun.star.frame.XComponentLoader;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;

/**
 * Can be used to interact/create different documents
 * 
 * @author Eelke van Turnhout <eelketurnhout3@gmail.com>
 * @version 1.0
 */
public class Desktop extends BaseConnection
{
    protected XComponentLoader componentLoader;

    /**
     * Construct a Desktop object
     * 
     * @param connection
     *            The connection a running OpenOffice service
     */
    public Desktop(Connection connection)
    {
        super(connection);
        this.initialize();
    }

    /**
     * Get the XComponentLoader. This can be used to create/open documents
     * 
     * @return XComponentLoader
     */
    public XComponentLoader getComponentLoader()
    {
        return this.componentLoader;
    }

    /**
     * Initializes the Desktop object
     */
    private void initialize()
    {
        Object desktop = null;

        try {
            desktop = this.connection.getServiceFactory().createInstance("com.sun.star.frame.Desktop");
        } catch (Exception e) {
            e.printStackTrace();
        }

        this.componentLoader = (XComponentLoader) UnoRuntime.queryInterface(XComponentLoader.class, desktop);
    }
}
