package dev.evt.openoffice;

import com.sun.star.bridge.XUnoUrlResolver;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.uno.UnoRuntime;

/**
 * <h1>Connection</h1>
 * <p>
 * Create a connection to an OpenOffice service running in headless mode.
 * </p>
 *
 * @author Eelke van Turnhout
 */
public class Connection {
    static final String UNO_URL_RESLOVER = "com.sun.star.bridge.UnoUrlResolver";

    protected XComponentLoader componentLoader;
    protected XMultiServiceFactory serviceFactory;
    protected XUnoUrlResolver urlResolver;
    protected boolean connected;

    /**
     * Connect to an OpenOffice service running in headless mode.
     *
     * @param ConnectionConfig the configurations needed to connect
     * @throws Exception
     */
    public Connection(ConnectionConfig config) throws Exception {
        this.connect(config);
    }

    /**
     * Connect to OpenOffice.
     *
     * @throws Exception
     */
    public void connect(ConnectionConfig config) throws Exception {
        XMultiServiceFactory xMultiServiceFactory = Bootstrap.createSimpleServiceManager();
        Object urlResolverObject = xMultiServiceFactory.createInstance(Connection.UNO_URL_RESLOVER);
        this.urlResolver = (XUnoUrlResolver) UnoRuntime.queryInterface(XUnoUrlResolver.class, urlResolverObject);
        String connectionStr = "uno:socket,host=" + config.getHost() + ",port=" + config.getPort()
                + ";urp;StarOffice.ServiceManager";
        Object initialObject = this.urlResolver.resolve(connectionStr);

        this.serviceFactory = (XMultiServiceFactory) UnoRuntime.queryInterface(XMultiServiceFactory.class,
                initialObject);

        this.componentLoader = (XComponentLoader) UnoRuntime.queryInterface(XComponentLoader.class,
            this.serviceFactory.createInstance("com.sun.star.frame.Desktop")
        );

        this.connected = true;
    }

    /**
     * Get the service factory. This is used to create other interfaces for
     * interaction with OpenOffice.
     *
     * @return the XMultiServiceFactory
     */
    public XMultiServiceFactory getServiceFactory() {
        return this.serviceFactory;
    }

    public XUnoUrlResolver getUrlResolver() {
        return this.urlResolver;
    }

    public XComponentLoader getComponentLoader() {
        return this.componentLoader;
    }

    /**
     * Check if a connection exists.
     *
     * @return boolean connected
     */
    public boolean isConnected() {
        return this.connected;
    }
}
