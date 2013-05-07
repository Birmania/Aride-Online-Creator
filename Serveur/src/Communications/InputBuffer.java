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

package Communications;

import java.io.ByteArrayInputStream;
import java.io.DataInputStream;
import java.io.IOException;

import Miscs.MessageLogger;


public class InputBuffer {
	private DataInputStream buffer;
	
	public InputBuffer(byte buffer[])
	{
		ByteArrayInputStream packet_buffer = new ByteArrayInputStream(buffer);
		this.buffer = new DataInputStream(packet_buffer);
	}
	
	public int readInt()
	{
		int rval = 0;
		
		try {
			rval = Integer.reverseBytes(this.buffer.readInt());
		} catch (IOException e) {
			MessageLogger.getInstance().log(e);
		}
		
		return rval;
	}

	public short readShort()
	{
		short rval = 0;
		
		try {
			rval = Short.reverseBytes(this.buffer.readShort());
		} catch (IOException e) {
			MessageLogger.getInstance().log(e);
		}
		
		return rval;
	}
	
	public String readString()
	{
		// get the string size
		int stringSize = this.readShort(); // Cast in an int to call the other readString method
		
		return this.readString(stringSize);
	}
	
	public String readString(int stringSize)
	{
		String rval = "";
		
		try
		{
			// create a byte tab which will represent the bytes of the read string
			byte byteString[] = new byte[stringSize];
	
			// read the corresponding bytes
			this.buffer.read(byteString);
			
			// convert the byte into a string
			try
			{
				rval = new String(byteString);
			} catch (OutOfMemoryError e) {
				MessageLogger.getInstance().log("Erreur de lecture d'un string dans le buffer. Nombre de bytes lus : "+stringSize);
			}
		} catch (IOException e) {
			MessageLogger.getInstance().log(e);
		}
		return rval;
	}
	
	public byte readByte()
	{
		byte rval = 0;
		
		try {
			rval = this.buffer.readByte();
		} catch (IOException e) {
			MessageLogger.getInstance().log(e);
		}
		
		return rval;
	}
	
	public boolean readBoolean()
	{
		boolean rval = false;
		
		try {
			rval = this.buffer.readBoolean();
			// Read boolean only use one byte so we must skip a non-use byte to respect the flow
			this.buffer.skipBytes(1);
		} catch (IOException e) {
			MessageLogger.getInstance().log(e);
		}
		
		return rval;
	}
	
	public long readLong()
	{
		long rval = 0;
		
		try {
			rval = Long.reverseBytes(this.buffer.readLong());
		} catch (IOException e) {
			MessageLogger.getInstance().log(e);
		}
		
		return rval;
	}
	
	public float readFloat()
	{
		float rval = 0;
		
		try {
			rval = Float.intBitsToFloat(Integer.reverseBytes(this.buffer.readInt()));
		} catch (IOException e) {
			MessageLogger.getInstance().log(e);
		}
		
		return rval;
	}
	
	/*public float readFloat()
	{
		float rval = 0;
		
		try {
			rval = this.buffer.readFloat();
		} catch (IOException e) {
			MessageLogger.getInstance().log(e);
		}
		
		return rval;
	}*/
	
	/*float readFloat()
	   {
	   // get 4 unsigned raw byte components, accumulating to an int,
	   // then convert the 32-bit pattern to a float.
	   int accum = 0;
	   for ( int shiftBy=0; shiftBy<32; shiftBy+=8 )
	      {
	      accum |= ( readByte () & 0xff ) << shiftBy;
	      }
	   return Float.intBitsToFloat( accum );

	   // there is no such method as Float.reverseBytes( f );

	   }*/
	
	public void skipBytes(int n)
	{
		try {
			this.buffer.skipBytes(n);
		} catch (IOException e) {
			MessageLogger.getInstance().log(e);
		}
	}
}
