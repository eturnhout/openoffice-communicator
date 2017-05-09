package dev.evt.openoffice;

import com.sun.star.frame.XComponentLoader;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;

/**
 * <h1>BaseConnection</h1>
 * <p>
 * Sets and makes the connection available for use in other classes when
 * extending
 * </p>
 * 
 * @author Eelke van Turnhout <eelketurnhout3@gmail.com>
 * @version 1.0
 */
public class BaseConnection
{
    protected Connection connection;
    protected XComponentLoader componentLoader;

    /**
     * Constructs a BaseConnection object.
     * 
     * @param connection
     *            The connection to a running OpenOffice service.
     */
    public BaseConnection(Connection connection)
    {
        this.setConnection(connection);
        this.initializeComponentLoader();
    }

    /**
     * Set the connection to interact with OpenOffice.
     * 
     * @param Connection
     *            The connection to a running OpenOffice service.
     */
    public void setConnection(Connection connection)
    {
        this.connection = connection;
    }

    /**
     * Get the connection from this BaseConnection object.
     * 
     * @return the connection to a running OpenOffice service.
     */
    public Connection getConnection()
    {
        return this.connection;
    }

    /**
     * Get the XComponentLoader. This can be used to create/open documents.
     * 
     * @return the XComponentLoader.
     */
    public XComponentLoader getComponentLoader()
    {
        return this.componentLoader;
    }

    /**
     * Initializes the XComponentLoader object.
     */
    private void initializeComponentLoader()
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
