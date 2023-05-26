package dev.evt.openoffice;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;

public class ConnectionTest
{
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
