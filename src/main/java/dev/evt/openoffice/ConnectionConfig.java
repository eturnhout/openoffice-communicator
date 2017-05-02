package dev.evt.openoffice;

public class ConnectionConfig
{
    protected String host;
    protected int port;
    
    /**
     * Create the configurations needed to connect to OpenOffice
     * 
     * @param host The host to connect to
     * @param port The port that the OpenOffice service is listening on
     */
    public ConnectionConfig(String host, int port)
    {
        this.setHost(host);
        this.setPort(port);
    }
    
    /**
     * Set the host on which OpenOffice is running
     * 
     * @param String host
     * @return void
     */
    public void setHost(String host) 
    {
        this.host = host;
    }
    
    /**
     * Get the host on which OpenOffice is running on
     * 
     * @param void
     * @return String host
     */
    public String getHost()
    {
        return this.host;
    }
    
    /**
     * Set the port on which OpenOffice is listening to
     * 
     * @param port
     * @return void
     */
    public void setPort(int port)
    {
        this.port = port;
    }
    
    /**
     * Get the port on which OpenOffice is listening to
     * 
     * @param void
     * @return integer port
     */
    public int getPort()
    {
        return this.port;
    }
}
