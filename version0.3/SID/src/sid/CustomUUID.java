package sid;

import java.security.SecureRandom;

public class CustomUUID {
	private long twepoch = 1288834974657L;
	private final static long regionIdBits = 3L;
	private final static long workerIdBits = 10L;
	private final static long sequenceBits = 10L;
	private final static long maxRegionId = -1L ^ (-1L << regionIdBits);
	private final static long maxWorkerId = -1L ^ (-1L << workerIdBits);
	private final static long sequenceMask = -1L ^ (-1L << sequenceBits);
	private final static long workerIdShift = sequenceBits;
	private final static long regionIdShift = sequenceBits + workerIdBits;
	private final static long timestampLeftShift = sequenceBits + workerIdBits + regionIdBits;
	private static long lastTimestamp = -1L;
	private long sequence = 0L;
	private final long workerId;
	private final long regionId;

	public CustomUUID(long workerId, long regionId) {
		if (workerId > maxWorkerId || workerId < 0) {
			throw new IllegalArgumentException("worker Id can't be greater than %d or less than 0");
		}
		if (regionId > maxRegionId || regionId < 0) {
			throw new IllegalArgumentException("worker Id can't be greater than %d or less than 0");
		}
		this.workerId=workerId;
		this.regionId=0;
	}
	public long generate() {
		return this.nextId(false,0);
	}
	private synchronized long nextId(boolean isPadding, long busId)
	{
		long timestamp=timeGen();
		long paddingnum=regionId;
		if(isPadding) {
			paddingnum=busId;
		}
		if(timestamp<lastTimestamp)
		{
			try {
				throw new Exception("Clock moved backwards. Refusing to generate id for"+(lastTimestamp-timestamp)+"millseconds");
			}catch(Exception e)
			{
				e.printStackTrace();
			}
		}
		if(lastTimestamp==timestamp) {
			sequence=(sequence+1)&sequenceMask;
			if(sequence==0) {
				timestamp=tailNextMillis(lastTimestamp);
			}
		}else {
			sequence =new SecureRandom().nextInt(10);
		}
		return ((timestamp-twepoch)<<timestampLeftShift)|(paddingnum<<regionIdShift)|sequence;
	}
	private long tailNextMillis(final long lastTimeStamp) {
		long timeStamp=this.timeGen();
		while(timeStamp<=lastTimeStamp) {
			timeStamp=this.timeGen();
		}
		return timeStamp;
	}
	private long timeGen() {
		return System.currentTimeMillis();
	}
}
