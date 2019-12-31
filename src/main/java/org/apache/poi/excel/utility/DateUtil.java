package org.apache.poi.excel.utility;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.OffsetDateTime;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.time.temporal.TemporalAccessor;
import java.util.Date;

import org.apache.commons.lang.time.DateUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class DateUtil {
	private final static Logger log = LoggerFactory.getLogger(DateUtil.class);

	/** The Constant DATE_FORMAT_YYYY_MM_DD. */
	public static final String DATE_FORMAT_YYYY_MM_DD = "yyyy-MM-dd";

	/**
	 * As date.
	 * 
	 * @param localDate the local date
	 * @return the date
	 */
	public static Date asDate(LocalDate localDate) {
		return Date.from(localDate.atStartOfDay().atZone(ZoneId.of("America/New_York")).toInstant());
	}

	/**
	 * As date.
	 * 
	 * @param localDateTime the local date time
	 * @return the date
	 */
	public static Date asDate(LocalDateTime localDateTime) {
		return Date.from(localDateTime.atZone(ZoneId.of("America/New_York")).toInstant());
	}

	/**
	 * As date.
	 * 
	 * @param offsetDateTime the offset date time
	 * @return the date
	 */
	public static Date asDate(OffsetDateTime offsetDateTime) {
		return Date.from(offsetDateTime.toInstant());
	}

	/**
	 * As date.
	 * 
	 * @param zonedDateTime the zoned date time
	 * @return the date
	 */
	public static Date asDate(ZonedDateTime zonedDateTime) {
		return Date.from(zonedDateTime.toInstant());
	}

	/**
	 * Convert any partial Date-time format into a java.util.Date. If it is unable
	 * to understand the passed format, it just returns null
	 * 
	 * @param timestamp the timestamp
	 * @return the string
	 */
	public static Date parse(String timestamp) {
		Date parsedDate = null;
		if (timestamp != null && !timestamp.trim().equals("")) {
			// First lets see if it is a standard ISO format
			try {
				// The submission time contains zone name like [America/New_york], [UTC-5] etc.
				// This is redundant and not a standard that a parser can understand. So lets
				// get rid of that.
				// Eg. 2018-10-18T14:36:19.419-05:00[UTC-05:00] becomes
				// 2018-10-18T14:36:19.419-05:00
				timestamp = timestamp.replaceAll("\\[.*\\]", "");

				// The different versions of Zulu time format that I could think of. [] implies
				// optional. [XXX] implies 'Z'oned time like -03:00.
				DateTimeFormatter formatter = DateTimeFormatter
						.ofPattern("yyyy-MM-dd[[ ]['T']HH:mm[:ss][.SSSSSSSSS][.SSSSSS][.SSS][.S][XXX]]");

				// Lets try to parse it and see if it matches anything that we understand
				TemporalAccessor ta = formatter.parseBest(timestamp, OffsetDateTime::from, LocalDateTime::from,
						LocalDate::from);

				// Check if the parsed result was any known format
				if (ta instanceof ZonedDateTime) {
					// An offset was present, convert it to New York Time.
					parsedDate = asDate(ZonedDateTime.from(ta));
				} else if (ta instanceof OffsetDateTime) {
					// An offset was present, convert it to New York Time.
					parsedDate = asDate(OffsetDateTime.from(ta));
				} else if (ta instanceof LocalDateTime) {
					// No Offset. Use it as a time stamp.
					parsedDate = asDate(LocalDateTime.from(ta));
				} else if (ta instanceof LocalDate) {
					// No time. Use it as a Date.
					parsedDate = asDate(LocalDate.from(ta));
				} else {
					throw new Exception("Cannot understand date-time format");
				}
			} catch (Exception p) {
				log.warn("The date was not in the standard format. Trying other known date formats.");
				// Hmm... Things did not work. Lets try something else
				try {
					// All alternate time formats you can think of
					parsedDate = DateUtils.parseDate(timestamp,
							new String[] { "MM/dd/yyyy", "MM/dd/yy", "yyyy/MM/dd", "yy/MM/dd", "mmddyy", "ddmmyy",
									"MMM dd, yy", "MMM dd, yyyy", "EEE, d MMM yyyy HH:mm:ss Z" });
					;
				} catch (Exception pex) {
					log.error("Unable to parse the passed date: '" + timestamp + "' to any known format");
				}
			}
		}
		return parsedDate;
	}
}