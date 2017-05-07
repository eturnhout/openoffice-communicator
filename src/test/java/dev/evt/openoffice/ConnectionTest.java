package dev.evt.openoffice;

import static org.junit.Assert.assertEquals;

import org.junit.Test;

public class ConnectionTest
{
    public ConnectionTest()
    {
    }

    /**
     * Tests if a connection can be made to OpenOffice
     * 
     * @throws Exception
     */
    @Test
    public void testConnection() throws Exception
    {
        QuickConnect quickConnect = new QuickConnect();
        Connection connection = quickConnect.getConnection();
        assertEquals(true, connection.isConnected());
    }
}
