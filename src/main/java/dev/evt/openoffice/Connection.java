package dev.evt.openoffice;

import com.sun.star.bridge.XUnoUrlResolver;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.uno.UnoRuntime;

/**
 * <h1>Connection</h1>
 * <p>
 * Create a connection to an OpenOffice service running in headless mode.
 * </p>
 * 
 * @author Eelke van Turnhout <eelketurnhout3@gmail.com>
 * @version 1.0
 */
public class Connection
{
    static final String UNO_URL_RESLOVER = "com.sun.star.bridge.UnoUrlResolver";

    protected ConnectionConfig config;
    protected XMultiServiceFactory serviceFactory;
    protected XUnoUrlResolver urlResolver;
    protected boolean connected;

    /**
     * Connect to an OpenOffice service running in headless mode.
     * 
     * @param ConnectionConfig
     *            The configurations needed to connect.
     * @throws Exception
     */
    public Connection(ConnectionConfig config) throws Exception
    {
        this.setConfig(config);
        this.connect();
    }

    /**
     * Set the connection's configurations.
     * 
     * @param ConnectionConfig
     *            The configurations needed to connect.
     */
    public void setConfig(ConnectionConfig config)
    {
        this.config = config;
    }

    /**
     * Get the connection's configurations.
     * 
     * @return the configurations needed to connect.
     */
    public ConnectionConfig getConfig()
    {
        return this.config;
    }

    /**
     * Connect to OpenOffice.
     * 
     * @throws Exception
     */
    public void connect() throws Exception
    {
        XMultiServiceFactory xMultiServiceFactory = Bootstrap.createSimpleServiceManager();
        Object urlResolverObject = xMultiServiceFactory.createInstance(Connection.UNO_URL_RESLOVER);
        this.urlResolver = (XUnoUrlResolver) UnoRuntime.queryInterface(XUnoUrlResolver.class, urlResolverObject);
        String connectionStr = "uno:socket,host=" + this.config.getHost() + ",port=" + this.config.getPort() + ";urp;StarOffice.ServiceManager";
        Object initialObject = this.urlResolver.resolve(connectionStr);

        this.serviceFactory = (XMultiServiceFactory) UnoRuntime.queryInterface(XMultiServiceFactory.class, initialObject);
        this.connected = true;
    }

    /**
     * Get the service factory. This is used to create other interfaces for
     * interaction with OpenOffice.
     * 
     * @return the XMultiServiceFactory
     */
    public XMultiServiceFactory getServiceFactory()
    {
        return this.serviceFactory;
    }

    public XUnoUrlResolver getUrlResolver()
    {
        return this.urlResolver;
    }

    /**
     * Check if a connection exists.
     * 
     * @return boolean connected
     */
    public boolean isConnected()
    {
        return this.connected;
    }
}
