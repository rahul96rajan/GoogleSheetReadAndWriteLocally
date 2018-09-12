package com.practiseAPI.util;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;

public class ConfigReaderUtil {
	public static Properties prop;
	public static void initialiseProperties() {	
		try {
			prop = new Properties();
			FileInputStream in = new FileInputStream(System.getProperty("user.dir")
					+ "/src/main/java/com/practise/config/config.properties");
			prop.load(in);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	public static String getProperty(String value) {
		return prop.getProperty(value);
	}
}