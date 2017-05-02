package dev.evt.openoffice;

/**
 * Create the configurations needed to connect to OpenOffice
 * 
 * @author Eelke van Turnhout <eelketurnhout3@gmail.com>
 * @version 1.0
 */
public class ConnectionConfig
{
    protected String host;
    protected int port;

    /**
     * Construct a ConnectionConfig object
     * 
     * @param host
     *            The host to connect to
     * @param port
     *            The port that the OpenOffice service is listening on
     */
    public ConnectionConfig(String host, int port)
    {
        this.setHost(host);
        this.setPort(port);
    }

    /**
     * Set the host on which OpenOffice is running
     * 
     * @param host
     *            The host to connect to
     */
    public void setHost(String host)
    {
        this.host = host;
    }

    /**
     * Get the host on which OpenOffice is running on
     * 
     * @return host The host to connect to
     */
    public String getHost()
    {
        return this.host;
    }

    /**
     * Set the port on which OpenOffice is listening to
     * 
     * @param port
     *            The port that the OpenOffice service is listening on
     */
    public void setPort(int port)
    {
        this.port = port;
    }

    /**
     * Get the port on which OpenOffice is listening to
     * 
     * @return port The port that the OpenOffice service is listening on
     */
    public int getPort()
    {
        return this.port;
    }
}
