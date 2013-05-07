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

import Communications.OutputBuffer;
import Main.ItemSlot;
import Main.Position;

public class MapItem {
	public Position position;
	public ItemSlot itemSlot;
	
	public MapItem(Position position, ItemSlot itemSlot)
	{
		this.position = position;
		this.itemSlot = itemSlot;
	}
	
	public void writeInPacket(OutputBuffer packet)
	{
		packet.writeByte(this.position.getX());
		packet.writeByte(this.position.getY());
		packet.writeShort(this.itemSlot.getItemId());
		packet.writeShort(this.itemSlot.getItemVal());
		packet.writeShort(this.itemSlot.getItemDur());
	}
}
