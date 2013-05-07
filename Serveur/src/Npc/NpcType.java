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

import Annotations.DeserializeIgnore;
import Communications.BinaryFile;
import Communications.InputBuffer;

public class NpcType extends BinaryFile {
	@DeserializeIgnore
	public short id;
	
	public String name;
	
	public short sprite;
	public int spawnSecs;	// Temps avant respawn
	public byte behavior;	// Quêteur, combattant, etc...
	public byte range;	// Vision
	
	// Statistics
	public short strength;
	public short defense;
	public short dexterity;
	public short science;
	public short language;
	
	public int maxHp;
	public int experience;	// experience donnée
	
	public byte spawnType;	// Spawn le jour ou la nuit ou les deux
	
	public short attackSpeed; // Durée entre deux attaques en millisecondes
	
	public NpcItem items[];
	
	public boolean immortal;
	public boolean fly;
	
	public short spells[];
	
	public NpcType(short id, InputBuffer buffer)
	{
		this.id = id;
		
		this.deserialize(buffer);
	}
}
