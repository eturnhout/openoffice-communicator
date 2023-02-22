package dev.evt.openoffice;

/**
 * <h1>ConnectionConfig</h1>
 * <p>
 * Create the configurations needed to connect to OpenOffice.
 * </p>
 *
 * @author Eelke van Turnhout
 */
final public class ConnectionConfig {
    private String host;
    private int port;

    /**
     * Construct a ConnectionConfig object.
     *
     * @param host the host to connect to
     * @param port the port that the OpenOffice service is listening on
     */
    public ConnectionConfig(String host, int port) {
        this.host = host;
        this.port = port;
    }

    /**
     * Get the host on which OpenOffice is running on.
     *
     * @return host the host to connect to
     */
    public String getHost() {
        return this.host;
    }

    /**
     * Get the port on which OpenOffice is listening to.
     *
     * @return the port that the OpenOffice service is listening on
     */
    public int getPort() {
        return this.port;
    }
}
