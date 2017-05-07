package dev.evt.openoffice;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

/**
 * A way to read config files.
 */
public class ConfigReader
{
    public static final String CONFIG_FILE = "config.properties";
    protected Properties propertyLoader;

    /**
     * Construct a ConfigReader object
     * 
     * @param configFile
     *            Set this to null to use the default value "config.properties"
     * @throws IOException
     */
    public ConfigReader(String configFile) throws IOException
    {
        if (configFile == null) {
            configFile = CONFIG_FILE;
        }

        this.initPropLoader(configFile);
    }

    /**
     * Get the value of the specified configuration
     * 
     * @param config
     *            Name of the configuration to fetch the value of
     * @return the configuration's value
     */
    public String getConfig(String config)
    {
        return this.propertyLoader.getProperty(config);
    }

    protected void initPropLoader(String configFile) throws IOException
    {
        this.propertyLoader = new Properties();
        InputStream input = null;

        try {
            input = new FileInputStream(configFile);

            // load the properties file
            propertyLoader.load(input);
        } catch (IOException ex) {
            throw new IOException(
                    "Please create a config file with the name \"config.properties\" and values \"host=yourhost\" and \"port=yourport\" to start testing.");
        } finally {
            if (input != null) {
                try {
                    input.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

}
