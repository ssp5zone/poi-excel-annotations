package org.apache.poi.excel.utility;

import java.io.File;
import java.io.FileNotFoundException;
import java.lang.reflect.Type;
import java.util.List;
import java.util.Scanner;

import com.google.gson.Gson;
import com.google.gson.reflect.TypeToken;

/**
 * 
 * A local utility class that helps in testing.
 * 
 * @author ssp5zone
 */
public class JsonReader {

	// Where mocks are located
	private static final String inPath = "src/test/resources/input/mocks/";

	// The GSON library for parsing String to JSON to POJO
	private static final Gson gson = new Gson();

	/**
	 * 
	 * This function reads the passed file present in test directory and converts it
	 * into a {@link List} of Java Objects.
	 * 
	 * @param <T>
	 * @param fileName
	 * @param klass    The Type class of which the List is to be returned
	 * @return the List of Objects read from the said files
	 * @throws FileNotFoundException
	 */
	public static <T> List<T> read(String fileName, Class<T> klass) throws FileNotFoundException {
		Type type = TypeToken.getParameterized(List.class, klass).getType();
		File file = new File(inPath.concat(fileName));
		Scanner in = new Scanner(file);
		String json = in.nextLine();
		in.close();
		return gson.fromJson(json, type);
	}
}
