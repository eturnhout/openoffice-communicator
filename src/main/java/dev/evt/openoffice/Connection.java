package dev.evt.openoffice;

import com.sun.star.bridge.XUnoUrlResolver;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.uno.UnoRuntime;

/**
 * Create a connection to an OpenOffice service running in headless mode
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
     * Connect to an openoffice service running in headless mode
     * 
     * @param ConnectionConfig
     *            config
     * @throws Exception
     */
    public Connection(ConnectionConfig config) throws Exception
    {
        this.setConfig(config);
        this.connect();
    }

    /**
     * Set the connection configurations
     * 
     * @param ConnectionConfig
     *            config
     * @return void
     */
    public void setConfig(ConnectionConfig config)
    {
        this.config = config;
    }

    /**
     * Get the connection configurations
     * 
     * @param void
     * @return ConnectionConfig
     */
    public ConnectionConfig getConfig()
    {
        return this.config;
    }

    /**
     * Connect to OpenOffice
     * 
     * @param void
     * @return void
     * @throws Exception
     */
    public void connect() throws Exception
    {
        XMultiServiceFactory xMultiServiceFactory = Bootstrap.createSimpleServiceManager();
        Object urlResolverObject = xMultiServiceFactory.createInstance(Connection.UNO_URL_RESLOVER);
        this.urlResolver = (XUnoUrlResolver) UnoRuntime.queryInterface(XUnoUrlResolver.class, urlResolverObject);
        String connectionStr = "uno:socket,host=" + this.config.getHost() + ",port=" + this.config.getPort()
                + ";urp;StarOffice.ServiceManager";
        Object initialObject = this.urlResolver.resolve(connectionStr);

        this.serviceFactory = (XMultiServiceFactory) UnoRuntime.queryInterface(XMultiServiceFactory.class,
                initialObject);
        this.connected = true;
    }

    /**
     * Get the service factory. This is used to create other interfaces for
     * interacting with OpenOffice
     * 
     * @return XMultiServiceFactory
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
     * Check if a connection exists
     * 
     * @param void
     * @return boolean connected
     */
    public boolean isConnected()
    {
        return this.connected;
    }
}
