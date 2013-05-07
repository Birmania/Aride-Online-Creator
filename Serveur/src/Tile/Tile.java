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

package Tile;

import Communications.BinaryFile;
import Communications.OutputBuffer;

public class Tile extends BinaryFile{
	public int ground;
	public int mask1;
	public int m1Anim;
	public int mask2;
	public int m2Anim;
	public int mask3;
	public int m3Anim;
	public int fringe1;
	public int f1Anim;
	public int fringe2;
	public int f2Anim;
	public int fringe3;
	public int f3Anim;
	public byte type;
	public int datas[];
	public String strings[];
	public int light;
	
	public Tile()
	{
	}
	
	public void writeInBuffer(OutputBuffer buffer)
	{
		buffer.writeInt(this.ground);
		buffer.writeInt(this.mask1);
		buffer.writeInt(this.m1Anim);
		buffer.writeInt(this.mask2);
		buffer.writeInt(this.m2Anim);
		buffer.writeInt(this.mask3);
		buffer.writeInt(this.m3Anim);
		buffer.writeInt(this.fringe1);
		buffer.writeInt(this.f1Anim);
		buffer.writeInt(this.fringe2);
		buffer.writeInt(this.f2Anim);
		buffer.writeInt(this.fringe3);
		buffer.writeInt(this.f3Anim);
		buffer.writeByte(this.type);

		int nbDatas;
		if (this.datas == null) { nbDatas = 0; } else { nbDatas = this.datas.length; }
		buffer.writeInt(nbDatas);
		for (int i = 0 ; i < nbDatas ; i++)
		{
			buffer.writeInt(this.datas[i]);
		}
		
		int nbStrings;
		if (this.strings == null) { nbStrings = 0; } else { nbStrings = this.strings.length; }
		buffer.writeInt(nbStrings);
		for (int i = 0 ; i < nbStrings ; i++)
		{
			buffer.writeString(this.strings[i]);
		}
		
		buffer.writeInt(this.light);
	}
}
