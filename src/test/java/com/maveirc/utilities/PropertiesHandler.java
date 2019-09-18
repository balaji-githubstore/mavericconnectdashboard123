package com.maveirc.utilities;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;

public class PropertiesHandler {
	/**
	 * Gives you the complete properties details
	 * throws IOException
	 * ***/
	public static Properties readProperties(String fileDescription) throws IOException
	{
		FileInputStream file = new FileInputStream(fileDescription);
		Properties prop = new Properties();
		prop.load(file);
		return prop;
	}
	
	public static String getValue(String fileDescription,String key) throws IOException
	{
		FileInputStream file = new FileInputStream(fileDescription);
		Properties prop = new Properties();
		prop.load(file);
	    String value=prop.getProperty(key);
		return value;
	}
	
	
}
