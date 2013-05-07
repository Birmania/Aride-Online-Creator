/*	Copyright (C) 2013  BRULTET Antoine
	
	This file is part of Aride Online Creator.

    Aride Online Creator is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    Aride Online Creator is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with Aride Online Creator.  If not, see <http://www.gnu.org/licenses/>.
*/

package Miscs;

import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;


public class TypeTools {
	
	public static final byte[] intToByteArray(int value) {
        return new byte[] {
                (byte)(value >>> 24),
                (byte)(value >>> 16),
                (byte)(value >>> 8),
                (byte)value};
	}
	
	public static final String byteArrayToMD5(byte array[])
	{
		// Determine the MD5 code of the map
		StringBuffer sb = new StringBuffer();
		try {
			byte[] md5 = MessageDigest.getInstance("MD5").digest(array);
			for (int j = 0; j < md5.length ; j++)
			{
				sb.append(Integer.toHexString((md5[j] & 0xFF) | 0x100).substring(1,3));
			}
		} catch (NoSuchAlgorithmException e) {
			MessageLogger.getInstance().log(e);
		}
		
		return sb.toString();
	}
	
	public static <T> List<T> substract(List<T> list1, List<T> list2)
	{
		List<T> result = new ArrayList<T>();
		Set<T> set2 = new HashSet<T>(list2);
		for (T t1 : list1)
		{
			if (!set2.contains(t1))
			{
				result.add(t1);
			}
		}
		return result;
	}
	
	public static String join(Object[] array, String glue)
	{
		StringBuilder rval = new StringBuilder();

		if (array.length > 0)
		{
			for (Object current : array)
			{
				rval.append(current.toString());
				rval.append(glue);
			}
		}
		
		return rval.toString();
	}
}