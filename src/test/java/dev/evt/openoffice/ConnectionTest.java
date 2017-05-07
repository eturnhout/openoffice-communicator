package dev.evt.openoffice;

import static org.junit.Assert.assertEquals;

import java.io.IOException;

import org.junit.Test;

public class ConnectionTest
{
    protected ConfigReader configReader;

    /**
     * Create the test case
     *
     * @param testName
     *            name of the test case
     * @throws IOException
     */
    public ConnectionTest() throws IOException
    {
        this.configReader = new ConfigReader(null);
    }

    /**
     * Tests if a connection can be made to OpenOffice
     * 
     * @throws Exception
     */
    @Test
    public void testConnection() throws Exception
    {
        String host = this.configReader.getConfig("host");
        int port = Integer.parseInt(this.configReader.getConfig("port"));
        ConnectionConfig config = new ConnectionConfig(host, port);
        Connection connection = new Connection(config);

        assertEquals(true, connection.isConnected());
    }
}
