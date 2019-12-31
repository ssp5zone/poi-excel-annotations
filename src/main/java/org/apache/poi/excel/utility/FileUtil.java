package org.apache.poi.excel.utility;

import java.io.File;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class FileUtil {
	private final static Logger log = LoggerFactory.getLogger(FileUtil.class);
	private static final long HR_PER_DAYS = 24;
	private static final long MIN_PER_HR = 60;
	private static final long SEC_PER_MIN = 60;
	private static final long MILSEC_PER_SEC = 1000;

	/**
	 * deleteFileOlder removes files in the given that are older than the given
	 * number of days
	 *
	 * @param path     Path of the files
	 * @param noOfDays Days that is the age of file exceeds to be deleted
	 * @return Count of files that were deleted
	 */
	public static int deleteFileOlderThanDays(String path, int noOfDays) {
		int deleteCount = 0;
		try {
			if (path == null || noOfDays < 0) {
				return 0;
			}
			File folder = new File(path);
			File[] listOfFiles = folder.listFiles();
			long span = noOfDays * HR_PER_DAYS * MIN_PER_HR * SEC_PER_MIN * MILSEC_PER_SEC;
			for (int i = 0; i < listOfFiles.length; i++) {
				if (listOfFiles[i].isFile()) {
					long diff = System.currentTimeMillis() - listOfFiles[i].lastModified();
					if (diff > span) {
						listOfFiles[i].delete();
						deleteCount++;
					}
				}
			}
		} catch (Exception ex) {

			log.error("Unable to delete files from the mentioned path: " + path);
		}
		return deleteCount;
	}
}
