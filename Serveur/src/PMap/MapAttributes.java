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

package PMap;
import Communications.BinaryFile;
import Communications.OutputBuffer;
import Npc.NpcAttributes;
import Tile.Tile;


public class MapAttributes extends BinaryFile{
	public String name;
	public byte mapType;
	public String music;
	public short respawnMap;
	public byte respawnX;
	public byte respawnY;
	public boolean indoors;
	public Tile tiles[][]; // La taille doit devenir variable en fonction des différentes maps
	public NpcAttributes npcs[];
	public String panoInf;
	public byte tranInf;
	public String panoSup;
	public byte tranSup;
	public short fog;
	public byte fogAlpha;
	public MapBorder borders[];
	public byte area;
	
	public MapAttributes()
	{
		
	}
	
	public void writeInBuffer(OutputBuffer buffer)
	{
		buffer.writeString(this.name);
		buffer.writeByte(this.mapType);
		//buffer.writeShort(this.upDestination);
		//buffer.writeShort(this.downDestination);
		//buffer.writeShort(this.leftDestination);
		//buffer.writeShort(this.rightDestination);
		buffer.writeString(this.music);
		buffer.writeShort(this.respawnMap);
		buffer.writeShort(this.respawnMap);
		buffer.writeByte(this.respawnX);
		buffer.writeByte(this.respawnY);
		buffer.writeBoolean(this.indoors);

		buffer.writeShort((short)this.tiles.length);
		buffer.writeShort((short)this.tiles[0].length);
		for (int x = 0 ; x < this.tiles.length ; x++)
		{
			for (int y = 0 ; y < this.tiles[0].length ; y++)
			{
				this.tiles[x][y].writeInBuffer(buffer);
			}
		}
		
		int nbNpcs;
		if (this.npcs == null) { nbNpcs = 0; } else { nbNpcs = this.npcs.length; }
		buffer.writeInt(nbNpcs);
		for (int i = 0 ; i < nbNpcs ; i++)
		{
			this.npcs[i].writeInBuffer(buffer);
		}
		
		buffer.writeString(this.panoInf);
		buffer.writeByte(this.tranInf);
		buffer.writeString(this.panoSup);
		buffer.writeByte(this.tranSup);
		buffer.writeShort(this.fog);
		buffer.writeByte(this.fogAlpha);
		
		int nbBorders;
		if (this.borders == null) { nbBorders = 0;} else { nbBorders = this.borders.length;}
		buffer.writeInt(nbBorders);
		for (int i = 0 ; i < nbBorders ; i++)
		{
			this.borders[i].writeInBuffer(buffer);
		}
		
		buffer.writeByte(this.area);
	}
}
