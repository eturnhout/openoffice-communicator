package dev.evt.openoffice;

/**
 * Sets and makes the connection available for use in other classes when extending
 *  
 * @author Eelke van Turnhout <eelketurnhout3@gmail.com>
 * @version 1.0
 */
public class BaseConnection
{
    protected Connection connection;
    
    /**
     * Constructs a BaseConnection object
     *
     * @param connection The connection to a running OpenOffice service
     */
    public BaseConnection(Connection connection)
    {
        this.setConnection(connection);
    }
    
    /**
     * Set the connection to interact with OpenOffice
     * 
     * @param Connection connection
     * @return void
     */
    public void setConnection(Connection connection)
    {
        this.connection = connection;
    }
    
    /**
     * Get the connection from this FileAccess object
     * 
     * @param void
     * @return Connection connection
     */
    public Connection getConnection()
    {
        return this.connection;
    }
}
