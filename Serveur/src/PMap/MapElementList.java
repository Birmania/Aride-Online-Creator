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

import java.util.ArrayList;
import java.util.Iterator;

import Interfaces.IKillable;
import Tile.TileBlock;

public class MapElementList {
	private ArrayList<MapElement> elements;
	
	public MapElementList(){
		this.elements = new  ArrayList<MapElement>();
	}
	
	public MapWalkable getMapWalkable()
	{
		synchronized(this.elements)
		{
			MapWalkable rval = null;
			Iterator<MapElement> ite = this.elements.iterator();
			
			while (ite.hasNext() && rval == null)
			{
				MapElement current = ite.next();
				if (current instanceof MapWalkable)
				{
					rval = (MapWalkable)current;
				}
			}
			
			return rval;
		}
	}
	
	public ArrayList<IKillable> getIKillable()
	{
		synchronized(this.elements)
		{
			ArrayList<IKillable> rval = new ArrayList<IKillable>();
			Iterator<MapElement> ite = this.elements.iterator();
			
			while  (ite.hasNext())
			{
				MapElement current = ite.next();
				if (current instanceof IKillable)
				{
					rval.add((IKillable)current);
					//rval = (IKillable)current;
				}
			}
			
			return rval;
		}
	}
	
	public ArrayList<MapMissile> getMapMissile()
	{
		synchronized(this.elements)
		{
			ArrayList<MapMissile> rval = new ArrayList<MapMissile>();
			Iterator<MapElement> ite = this.elements.iterator();
			
			while  (ite.hasNext())
			{
				MapElement current = ite.next();
				if (current instanceof MapMissile)
				{
					rval.add((MapMissile)current);
					//rval = (IKillable)current;
				}
			}
			
			return rval;
		}
	}
	
	public boolean isTileBlock()
	{
		synchronized(this.elements)
		{
			boolean rval = false;
			Iterator<MapElement> ite = this.elements.iterator();
			
			while (ite.hasNext() && rval == false)
			{
				if (ite.next() instanceof TileBlock)
				{
					rval = true;
				}
			}
			
			return rval;
		}
	}
	
	public boolean isTraversableBy(MapMovable element)
	{
		//WARNING : Ne pas laisser d'accès à l'écran dans cette fonction !
		boolean rval = true;
		
		if (this.getMapWalkable() != null || this.isTileBlock())
		{
			rval = false;
		}
		
		return rval;
	}
	
	public int size()
	{
		return this.elements.size();
	}
	
	public void add(MapElement element)
	{
		synchronized(this.elements)
		{
			this.elements.add(element);
		}
	}
	
	public void add(MapElementList elementList)
	{
		synchronized(this.elements)
		{
			synchronized(elementList.elements)
			{
				for (MapElement currentElement : elementList.elements)
				{
					this.add(currentElement);
				}
			}
		}
	}
	
	public void remove(MapElement element)
	{
		synchronized(this.elements)
		{
			this.elements.remove(element);
		}
	}
}
