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
 * @author Eelke van Turnhout
 */
public class BaseConnection {
    protected Connection connection;
    protected XComponentLoader componentLoader;

    /**
     * Constructs a BaseConnection object.
     *
     * @param connection
     *                   The connection to a running OpenOffice service.
     *
     * @throws Exception
     */
    public BaseConnection(Connection connection) throws Exception {
        this.connection = connection;

        this.componentLoader = (XComponentLoader) UnoRuntime.queryInterface(XComponentLoader.class,
                connection.getServiceFactory().createInstance("com.sun.star.frame.Desktop"));
    }
}
