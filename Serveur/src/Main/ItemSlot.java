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

package Main;

public class ItemSlot {
	private short itemId;
	private short itemVal;
	private short itemDur;
	
	public ItemSlot(short id, short val, short dur)
	{
		this.setItemId(id);
		this.itemVal = val;
		this.setItemDur(dur);
	}
	
	public ItemSlot(ItemSlot itemSlot)
	{
		this.setItemId(itemSlot.getItemId());
		this.addItemVal(itemSlot.getItemVal());
		this.setItemDur(itemSlot.getItemDur());
	}
	
	public void setItemId(short id)
	{
		this.itemId = id;
		//this.itemVal = -1;
		this.itemVal = 0;
		this.setItemDur((short)-1);
	}
	
	public short getItemId()
	{
		return this.itemId;
	}
	
	public void addItemVal(short val)
	{
		this.itemVal += val;
	}
	
	public void reduceItemVal(short val)
	{
		this.itemVal -= val;
		if (this.itemVal <= 0)
		{
			this.clear();
		}
	}
	
	public short getItemVal()
	{
		return this.itemVal;
	}
	
	public void setItemDur(short dur)
	{
		this.itemDur = dur;
	}
	
	public short getItemDur()
	{
		return this.itemDur;
	}
	
	public void clear()
	{
		this.itemDur = -1;
		this.itemVal = -1;
		this.itemId = -1;
	}
}
