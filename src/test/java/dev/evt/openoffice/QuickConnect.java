package dev.evt.openoffice;

public class QuickConnect
{
    protected ConfigReader configReader;
    protected Connection connection;

    public QuickConnect() throws Exception
    {
        this.configReader = new ConfigReader(null);
        this.initialize();
    }

    public Connection getConnection()
    {
        return this.connection;
    }

    public ConfigReader getConfigReader()
    {
        return this.configReader;
    }

    private void initialize() throws Exception
    {
        String host = this.configReader.getConfig("host");
        int port = Integer.parseInt(this.configReader.getConfig("port"));
        ConnectionConfig config = new ConnectionConfig(host, port);
        this.connection = new Connection(config);
    }
}
