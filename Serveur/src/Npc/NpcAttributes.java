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

package Npc;
import Communications.BinaryFile;
import Communications.OutputBuffer;


public class NpcAttributes extends BinaryFile {
	public short id;	// id of the npc type
	public byte x[];
	public byte y[];
	public byte dir;
	public boolean hasardSpawn;
	public byte movementType;
	
	
	public void writeInBuffer(OutputBuffer buffer)
	{
		buffer.writeShort(this.id);
		buffer.writeInt(this.x.length);
		for (int i = 0 ; i < this.x.length ; i++)
		{
			buffer.writeByte(this.x[i]);
		}
		
		buffer.writeInt(this.y.length);
		for (int i = 0 ; i < this.y.length ; i++)
		{
			buffer.writeByte(this.y[i]);
		}
		
		buffer.writeByte(this.dir);
		buffer.writeBoolean(this.hasardSpawn);
		buffer.writeByte(this.movementType);
	}
}
