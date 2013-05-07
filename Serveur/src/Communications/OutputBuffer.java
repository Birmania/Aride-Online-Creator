package Communications;



import java.io.ByteArrayOutputStream;
import java.io.DataOutputStream;
import java.io.IOException;

import Miscs.MessageLogger;

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

public class OutputBuffer {
	private ByteArrayOutputStream packet;
	private DataOutputStream writer;
	
	private String packetType;
	public String test()
	{
		return this.packetType;
	}
	
	public OutputBuffer(String packetType) {
		this.packetType = packetType;
		this.packet = new ByteArrayOutputStream();
		this.writer = new DataOutputStream(this.packet);
		if (!packetType.equals(""))
		{
			try {
				this.writeInt(HandleData.getInstance().serverPackets.get(packetType));
			} catch (NullPointerException m) {
				MessageLogger.getInstance().log("Le paquet "+packetType+" n'existe pas dans le tableau des paquets disponibles.");
			}
		}
	}
	
	public void writeInt(int value) {
		try {
			this.writer.writeInt(Integer.reverseBytes(value));
		} catch (IOException e) {
			MessageLogger.getInstance().log(e);
		}
	}
	
	public void writeShort (short value) {
		try {
			this.writer.writeShort(Short.reverseBytes(value));
		} catch (IOException e) {
			MessageLogger.getInstance().log(e);
		}
	}
	
	public void writeString(String value)
	{
		byte byteString[] = value.getBytes();
		
		try {
			this.writeInt(byteString.length);
			this.writer.write(byteString);
		} catch (IOException e) {
			MessageLogger.getInstance().log(e);
		}
	}
	
	public void writeByte(byte value)
	{		
		try {
			byte v[] = new byte[1];
			v[0] = value;
			this.writer.write(v, 0, 1);
		} catch (IOException e) {
			MessageLogger.getInstance().log(e);
		}
	}
	
	public void writeBoolean(boolean value)
	{
		try {
			this.writer.writeBoolean(value);
		} catch (IOException e) {
			MessageLogger.getInstance().log(e);
		}
	}
	
	/*public void writeFloat(Float value)
	{	
		try {
			this.writer.writeFloat(value);
		} catch (IOException e) {
			MessageLogger.getInstance().log(e);
		}
	}*/
	
	public byte[] toByteArray()
	{
		return this.packet.toByteArray();
	}
	
	public void writePacket(OutputBuffer packet)
	{
		try {
			this.writer.write(packet.toByteArray());
		} catch (IOException e) {
			MessageLogger.getInstance().log(e);
		}
	}
}