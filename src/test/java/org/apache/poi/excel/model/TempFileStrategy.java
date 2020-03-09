package org.apache.poi.excel.model;

import java.io.File;
import java.io.IOException;
import java.nio.file.Paths;

import org.apache.poi.util.TempFileCreationStrategy;

/**
 * This is only needed to fix a POI Temp directory bug during running the test
 * cases. Refer:
 * https://stackoverflow.com/questions/29285076/java-apache-poi-sxssfworkbook-unable-to-create-sheets
 * <br>
 * This is alternate solution that I have tried based on the source code.
 * 
 * @author ssp5zone
 */
public class TempFileStrategy implements TempFileCreationStrategy {

	private static final String tempDir = "build/output/temp";

	/**
	 * @see org.apache.poi.util.TempFileCreationStrategy#createTempFile(java.lang.String,
	 *      java.lang.String)
	 */
	@Override
	public File createTempFile(String prefix, String suffix) throws IOException {
		File file = new File(Paths.get(tempDir, prefix.concat(suffix)).toString());
		file.createNewFile();
		return file;
	}

	/**
	 * @see org.apache.poi.util.TempFileCreationStrategy#createTempDirectory(java.lang.String)
	 */
	@Override
	public File createTempDirectory(String prefix) throws IOException {
		File file = new File(Paths.get(tempDir, prefix).toString());
		file.mkdirs();
		return file;
	}

}
