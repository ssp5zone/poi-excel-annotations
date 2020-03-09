package org.apache.poi.excel;

import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.greaterThan;
import static org.junit.Assert.assertTrue;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.List;

import org.apache.poi.excel.model.ExcelAnnotated;
import org.apache.poi.excel.model.ExcelEdge;
import org.apache.poi.excel.model.ExcelNonAnnotated;
import org.apache.poi.excel.model.TempFileStrategy;
import org.apache.poi.excel.utility.JsonReader;
import org.apache.poi.util.TempFile;
import org.junit.BeforeClass;
import org.junit.Test;

public class ExcelWriterTest {

	private static List<ExcelNonAnnotated> nonAnnontatedPojo;
	private static List<ExcelAnnotated> annontatedPojo;
	private static List<ExcelEdge> edgePojo;

	private static final String outPath = "src/test/resources/output/";

	/**
	 * Creates a temp directory override for Apache POI. Only needed during testing.
	 * 
	 * @throws IOException
	 */
	@BeforeClass
	public static void onlyOnce() throws IOException {
		TempFileStrategy strategy = new TempFileStrategy();
		strategy.createTempDirectory("");
		TempFile.setTempFileCreationStrategy(strategy);
	}

	/**
	 * Initialized mocks needed for testing.
	 * 
	 * @throws FileNotFoundException
	 */
	@BeforeClass
	public static void initMocks() throws FileNotFoundException {
		nonAnnontatedPojo = JsonReader.read("NonAnnotated.json", ExcelNonAnnotated.class);
		annontatedPojo = JsonReader.read("Annotated.json", ExcelAnnotated.class);
		edgePojo = JsonReader.read("EdgeCases.json", ExcelEdge.class);
	}

	@Test
	public void testNonAnnontated() {
		String testFileName = "NonAnnotated.xlsx";
		ExcelWriter.write(outPath, testFileName, nonAnnontatedPojo);
		File file = new File(outPath.concat(testFileName));
		assertTrue(file.exists());
		assertThat(file.length(), greaterThan(0L));
	}

	@Test
	public void testAnnontated() {
		String testFileName = "Annotated.xlsx";
		ExcelWriter.write(outPath, testFileName, annontatedPojo);
		File file = new File(outPath.concat(testFileName));
		assertTrue(file.exists());
		assertThat(file.length(), greaterThan(0L));
	}

	@Test
	public void testEdge() {
		String testFileName = "EdgeCases.xlsx";
		ExcelWriter.write(outPath, testFileName, edgePojo);
		File file = new File(outPath.concat(testFileName));
		assertTrue(file.exists());
		assertThat(file.length(), greaterThan(0L));
	}

	@Test
	public void testMixed() {
		String testFileName = "MultiSheets.xlsx";
		ExcelWriter.write(outPath, testFileName, annontatedPojo, nonAnnontatedPojo, edgePojo);
		File file = new File(outPath.concat(testFileName));
		assertTrue(file.exists());
		assertThat(file.length(), greaterThan(0L));
	}

}
